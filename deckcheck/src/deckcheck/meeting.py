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
from deckcheck.changeset.engine import Invalid, Problem, load, parse
from deckcheck.cli import write_atomic
from deckcheck.model import DeckError

OK, PROBLEMS, USAGE, NO_CLAUDE = 0, 1, 2, 3
SKILL = ".claude/skills/process-meeting/SKILL.md"
DOC = "docs/changeset.md"
DOCS = (SKILL, DOC)
ENGINE = "deckcheck"
OUT_ROOTS = ("artifacts", "private")
NOT_LOGGED_IN = "claude CLI not found or not logged in; run `claude` once to log in"
DECK = "before.pptx"
CHANGESET = "changeset.json"
COMMANDS = (
    f"uv run --project {ENGINE} changeset outline {DECK} --json outline.json",
    f"uv run --project {ENGINE} changeset validate {CHANGESET}",
    f"uv run --project {ENGINE} changeset execute {CHANGESET} --out executed.pptx --review review.json",
    f"uv run --project {ENGINE} deckcheck check executed.pptx --out check",
)
TOOLS = "Read,Write,Grep,Glob,Bash"
ALLOWED = ("Edit(./**)", *(f"Bash({c})" for c in COMMANDS))


class Refused(ValueError):
    pass


@dataclass(frozen=True)
class Job:
    transcript: str
    data: tuple[str, ...]
    deck: str
    out: str

    @property
    def stage(self) -> str:
        return f"{self.out}/stage"

    @property
    def changeset(self) -> str:
        return f"{self.out}/{CHANGESET}"

    @property
    def claude_output(self) -> str:
        return f"{self.out}/claude.json"


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
    source = _inside(root, deck if deck is not None else root / meeting_dir / DECK, "the deck")
    if not (root / source).is_file():
        raise Refused(f"no deck at {source}; pass --deck")
    out_dir = _inside(root, out, "--out")
    if out_dir.split("/")[0] not in OUT_ROOTS or "/" not in out_dir:
        raise Refused(f"--out {out_dir} must sit inside artifacts/ or private/, which git ignores, so client files stay out of git")
    data = tuple(p.relative_to(root).as_posix() for p in sorted((root / meeting_dir / "data").glob("*.csv")))
    run = Job(transcript=transcript, data=data, deck=source, out=out_dir)
    if any(_within(p, run.stage) for p in (meeting_dir, source)):
        raise Refused(f"{run.stage} is replaced on every run, so the meeting and the deck cannot sit inside it")
    return run


def _inside(root: Path, path: Path, what: str) -> str:
    full = (root / path).resolve()
    if not full.is_relative_to(root):
        raise Refused(f"{what} {path} is outside the repo root {root}")
    return full.relative_to(root).as_posix()


def _within(path: str, folder: str) -> bool:
    return path == folder or path.startswith(folder + "/")


def stage(root: Path, job: Job) -> Path:
    folder = root / job.stage
    if folder.exists():
        shutil.rmtree(folder)
    copies = {
        posixpath.basename(job.transcript): job.transcript,
        **{f"data/{posixpath.basename(d)}": d for d in job.data},
        DECK: job.deck,
        **{d: d for d in DOCS},
    }
    for name, source in copies.items():
        (folder / name).parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(root / source, folder / name)
    (folder / ENGINE).symlink_to(root / ENGINE, target_is_directory=True)
    return folder


def prompt(job: Job) -> str:
    data = ", ".join(f"data/{posixpath.basename(d)}" for d in job.data) or "none"
    commands = "\n".join(f"- `{c}`" for c in COMMANDS)
    return f"""Follow the skill at {SKILL} to turn this meeting into a ChangeSet for its deck.

Your working folder was staged for this run. It holds:
- the meeting: {posixpath.basename(job.transcript)}
- data files: {data}
- the source deck: {DECK}
- the skill, and {DOC}, which describes every op and field

Write the ChangeSet to {CHANGESET}.

Read files with Read, Grep, and Glob. Write files with Write. Every file you need is in the working folder, and you can read and write nothing outside it.

These are the only commands you can run. Run each exactly as written, from the working folder:
{commands}

The user ran `meeting process --shareable`, which says this deck and meeting may be shared with an AI service under their firm's policy. Do not ask again.

Nobody can answer questions during this run. Put anything you cannot settle in the ChangeSet's flags.
"""


def argv(job: Job) -> list[str]:
    return [
        "claude",
        "-p",
        prompt(job),
        "--tools",
        TOOLS,
        "--restricted",
        "--strict-mcp-config",
        "--permission-mode",
        "dontAsk",
        "--allowedTools",
        *ALLOWED,
        "--output-format",
        "json",
        "--no-session-persistence",
    ]


# Every other tool reads source.path from the repo root, so the published ChangeSet names the deck there.
def publish(staged: Path, published: Path, deck: str) -> None:
    try:
        raw = staged.read_text(encoding="utf-8-sig")
    except (OSError, UnicodeDecodeError) as e:
        raise Invalid([Problem("changeset", f"cannot read it: {e}")]) from e
    named = parse(raw).source.path
    if posixpath.normpath(named) != DECK:
        raise Invalid([Problem("source.path", f"{named!r} is not the deck this meeting edits; name {DECK}")])
    doc = json.loads(raw)
    doc["source"]["path"] = deck
    write_atomic(published, (json.dumps(doc, indent=2, ensure_ascii=False) + "\n").encode())


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
    # A ChangeSet left by an earlier run would otherwise pass for this run's output.
    (root / run.changeset).unlink(missing_ok=True)
    folder = stage(root, run)
    claude = subprocess.run(argv(run), cwd=folder, stdin=subprocess.DEVNULL, stdout=subprocess.PIPE)
    write_atomic(root / run.claude_output, claude.stdout)
    if claude.returncode:
        print(f"error: claude -p exited {claude.returncode}", file=sys.stderr)
    staged = folder / CHANGESET
    if not staged.is_file():
        print(f"error: claude wrote no ChangeSet at {run.stage}/{CHANGESET}", file=sys.stderr)
        return PROBLEMS
    try:
        publish(staged, root / run.changeset, run.deck)
    except Invalid as e:
        return print_invalid(Path(f"{run.stage}/{CHANGESET}"), e.problems)
    try:
        return cmd_validate(load(Path(run.changeset)))
    except Invalid as e:
        return print_invalid(Path(run.changeset), e.problems)
    except (DeckError, OSError) as e:
        print(f"error: {e}", file=sys.stderr)
        return USAGE
