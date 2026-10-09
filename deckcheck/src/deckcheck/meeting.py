from __future__ import annotations

import argparse
import json
import posixpath
import shutil
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path

from deckcheck.changeset.cli import cmd_validate, print_invalid
from deckcheck.changeset.engine import Invalid, Problem, load
from deckcheck.cli import write_atomic
from deckcheck.model import DeckError

OK, PROBLEMS, USAGE, NO_CLAUDE = 0, 1, 2, 3
SKILL = ".claude/skills/process-meeting/SKILL.md"
OUT_ROOTS = ("artifacts", "private")
NOT_LOGGED_IN = "claude CLI not found or not logged in; run `claude` once to log in"
UV = "uv run --project deckcheck"


class Refused(ValueError):
    pass


@dataclass(frozen=True)
class Job:
    meeting: str
    transcript: str
    data: tuple[str, ...]
    deck: str
    out: str

    @property
    def changeset(self) -> str:
        return f"{self.out}/changeset.json"

    @property
    def executed(self) -> str:
        return f"{self.out}/executed.pptx"

    @property
    def check(self) -> str:
        return f"{self.out}/check"

    @property
    def outline(self) -> str:
        return f"{self.out}/outline.json"

    @property
    def review(self) -> str:
        return f"{self.out}/review.json"

    @property
    def claude_output(self) -> str:
        return f"{self.out}/claude.json"

    @property
    def commands(self) -> tuple[str, ...]:
        return (
            f"{UV} changeset outline {self.deck} --json {self.outline}",
            f"{UV} changeset validate {self.changeset}",
            f"{UV} changeset execute {self.changeset} --out {self.executed} --review {self.review}",
            f"{UV} deckcheck check {self.executed} --out {self.check}",
        )


def job(root: Path, meeting: Path, out: Path, deck: Path | None, shareable: bool) -> Job:
    if not (root / SKILL).is_file():
        raise Refused(f"no {SKILL} here; run meeting process from the repo root")
    if not shareable:
        raise Refused(
            "pass --shareable only if this deck and meeting may be shared with an AI service under your firm's policy "
            "(CLAUDE.md); the run sends them to Claude"
        )
    meeting_dir = _inside(root, meeting, "the meeting folder")
    if not (root / meeting_dir).is_dir():
        raise Refused(f"{meeting_dir} is not a folder")
    transcript = next((f"{meeting_dir}/{n}" for n in ("transcript.md", "notes.md") if (root / meeting_dir / n).is_file()), None)
    if transcript is None:
        raise Refused(f"{meeting_dir} holds neither transcript.md nor notes.md")
    source = _inside(root, deck if deck is not None else root / meeting_dir / "before.pptx", "the deck")
    if not (root / source).is_file():
        raise Refused(f"no deck at {source}; pass --deck")
    out_dir = _inside(root, out, "--out")
    if out_dir.split("/")[0] not in OUT_ROOTS or "/" not in out_dir:
        raise Refused(f"--out {out_dir} must sit inside artifacts/ or private/, which git ignores, so client files stay out of git")
    data = tuple(p.relative_to(root).as_posix() for p in sorted((root / meeting_dir / "data").glob("*.csv")))
    return Job(meeting=meeting_dir, transcript=transcript, data=data, deck=source, out=out_dir)


def _inside(root: Path, path: Path, what: str) -> str:
    full = (root / path).resolve()
    if not full.is_relative_to(root):
        raise Refused(f"{what} {path} is outside the repo root {root}")
    return full.relative_to(root).as_posix()


def prompt(job: Job) -> str:
    commands = "\n".join(f"- `{c}`" for c in job.commands)
    return f"""Follow the skill at {SKILL} to turn this meeting into a ChangeSet for its deck.

- Meeting folder: {job.meeting}
- Transcript: {job.transcript}
- Data files: {", ".join(job.data) or "none"}
- Source deck: {job.deck}
- Working folder: {job.out}
- ChangeSet: write it to {job.changeset}

These are the only commands you can run. Run each exactly as written:
{commands}

The user ran `meeting process --shareable`, which says this deck and meeting may be shared with an AI service under their firm's policy. Do not ask again.

Nobody can answer questions during this run. Put anything you cannot settle in the ChangeSet's flags.
"""


def argv(job: Job) -> list[str]:
    allowed = [
        f"Read(./{job.meeting}/**)",
        f"Read(./{posixpath.dirname(SKILL)}/**)",
        "Read(./docs/changeset.md)",
        f"Read(./{job.out}/**)",
        f"Write(./{job.out}/**)",
        f"Edit(./{job.out}/**)",
        *(f"Bash({c})" for c in job.commands),
    ]
    return [
        "claude",
        "-p",
        prompt(job),
        "--tools",
        "Read,Write,Edit,Bash",
        "--permission-mode",
        "dontAsk",
        "--allowedTools",
        *allowed,
        "--output-format",
        "json",
        "--no-session-persistence",
    ]


def source_problems(named: str, job: Job) -> list[Problem]:
    if posixpath.normpath(named) == job.deck:
        return []
    return [Problem("source.path", f"{named!r} is not the deck this meeting edits; name {job.deck}")]


def logged_in() -> bool:
    if shutil.which("claude") is None:
        return False
    status = subprocess.run(["claude", "auth", "status", "--json"], capture_output=True, text=True)
    try:
        return status.returncode == 0 and json.loads(status.stdout).get("loggedIn") is True
    except (json.JSONDecodeError, AttributeError):
        return False


def main(argv_: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="meeting", description="Turn a meeting into a ChangeSet for its deck with Claude Code.")
    sub = parser.add_subparsers(dest="command", required=True)
    process = sub.add_parser("process", help="run the process-meeting skill headless, then validate the ChangeSet it wrote")
    process.add_argument("meeting", type=Path, help="the meeting folder, holding transcript.md or notes.md and optional data/*.csv")
    process.add_argument("--out", type=Path, required=True, help="the working folder, inside artifacts/ or private/")
    process.add_argument("--deck", type=Path, help="the source deck; default <meeting>/before.pptx")
    process.add_argument("--shareable", action="store_true", help="this deck and meeting may be shared with an AI service under your firm's policy")
    args = parser.parse_args(argv_)
    root = Path.cwd().resolve()
    try:
        run = job(root, args.meeting, args.out, args.deck, args.shareable)
    except Refused as e:
        print(f"error: {e}", file=sys.stderr)
        return USAGE
    if not logged_in():
        print(f"error: {NOT_LOGGED_IN}", file=sys.stderr)
        return NO_CLAUDE
    changeset = root / run.changeset
    # A ChangeSet left by an earlier run would otherwise pass for this run's output.
    changeset.unlink(missing_ok=True)
    (root / run.out).mkdir(parents=True, exist_ok=True)
    claude = subprocess.run(argv(run), cwd=root, stdin=subprocess.DEVNULL, stdout=subprocess.PIPE)
    write_atomic(root / run.claude_output, claude.stdout)
    if claude.returncode:
        print(f"error: claude -p exited {claude.returncode}", file=sys.stderr)
    if not changeset.is_file():
        print(f"error: claude wrote no ChangeSet at {run.changeset}", file=sys.stderr)
        return PROBLEMS
    try:
        checked = load(Path(run.changeset))
        if problems := source_problems(checked.changeset.source.path, run):
            raise Invalid(problems)
        return cmd_validate(checked)
    except Invalid as e:
        return print_invalid(Path(run.changeset), e.problems)
    except (DeckError, OSError) as e:
        print(f"error: {e}", file=sys.stderr)
        return USAGE
