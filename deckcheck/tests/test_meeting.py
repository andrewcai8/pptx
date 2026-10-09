from __future__ import annotations

import json
import shutil
import sys
from pathlib import Path

import pytest
from test_changeset import CELL, REVENUE, build_deck, change, changeset

from deckcheck.meeting import main

DECK = "private/m1/before.pptx"
FORBIDDEN = (
    "private/m1/after.pptx",
    "private/m1/expected.yaml",
    "private/m1/changeset.json",
    "private/m1/build.py",
    "private/m1/notes.md",
    "private/m2/transcript.md",
    "CLAUDE.md",
)
FAKE = """#!{python}
import json, os, shutil, sys

args = sys.argv[1:]
if args[:2] == ["auth", "status"]:
    print(json.dumps({{"loggedIn": {logged_in}, "authMethod": "none"}}))
    sys.exit(0 if {logged_in} else 1)
cwd = os.path.realpath(os.getcwd())
roots = [cwd, *(os.path.realpath(args[i + 1]) for i, a in enumerate(args) if a == "--add-dir")]
rules = args[args.index("--allowedTools") + 1 : args.index("--output-format")]


def grant(path):
    real = os.path.realpath(path)
    if "--restricted" not in args:
        return "no --restricted, so user and project settings and read-only shell commands reach it"
    if any(real == r or real.startswith(r + os.sep) for r in roots):
        return "inside a working directory"
    shell = [r for r in rules if r.startswith("Bash(") and ("*" in r or os.path.relpath(real, cwd) in r or real in r)]
    return shell[0] if shell else None


files = sorted(
    os.path.relpath(os.path.join(d, n), cwd) + ("@" if os.path.islink(os.path.join(d, n)) else "")
    for d, dirs, names in os.walk(cwd)
    for n in names + [x for x in dirs if os.path.islink(os.path.join(d, x))]
)
reads = {{p: grant(os.path.join({root!r}, p)) for p in {forbidden!r}}}
with open({log!r}, "a") as log:
    log.write(json.dumps({{"cwd": os.path.relpath(cwd, {root!r}), "argv": args, "files": files, "reads": reads}}) + "\\n")
if {fixture!r}:
    shutil.copy({fixture!r}, "changeset.json")
print(json.dumps({{"type": "result", "result": "done"}}))
"""
PROMPT = """Follow the skill at .claude/skills/process-meeting/SKILL.md to turn this meeting into a ChangeSet for its deck.

Your working folder was staged for this run. It holds:
- the meeting: transcript.md
- data files: data/x.csv
- the source deck: before.pptx
- the skill, and docs/changeset.md, which describes every op and field

Write the ChangeSet to changeset.json.

Read files with Read, Grep, and Glob. Write files with Write. Every file you need is in the working folder, and you can read and write nothing outside it.

These are the only commands you can run. Run each exactly as written, from the working folder:
- `uv run --project deckcheck changeset outline before.pptx --json outline.json`
- `uv run --project deckcheck changeset validate changeset.json`
- `uv run --project deckcheck changeset execute changeset.json --out executed.pptx --review review.json`
- `uv run --project deckcheck deckcheck check executed.pptx --out check`

The user ran `meeting process --shareable`, which says this deck and meeting may be shared with an AI service under their firm's policy. Do not ask again.

Nobody can answer questions during this run. Put anything you cannot settle in the ChangeSet's flags.
"""
ARGV = [
    "-p",
    PROMPT,
    "--tools",
    "Read,Write,Grep,Glob,Bash",
    "--restricted",
    "--strict-mcp-config",
    "--permission-mode",
    "dontAsk",
    "--allowedTools",
    "Edit(./**)",
    "Bash(uv run --project deckcheck changeset outline before.pptx --json outline.json)",
    "Bash(uv run --project deckcheck changeset validate changeset.json)",
    "Bash(uv run --project deckcheck changeset execute changeset.json --out executed.pptx --review review.json)",
    "Bash(uv run --project deckcheck deckcheck check executed.pptx --out check)",
    "--output-format",
    "json",
    "--no-session-persistence",
]
STAGED = [
    ".claude/skills/process-meeting/SKILL.md",
    "before.pptx",
    "data/x.csv",
    "deckcheck@",
    "docs/changeset.md",
    "transcript.md",
]
ARGS = ["process", "private/m1", "--out", "artifacts/run1", "--shareable"]


@pytest.fixture
def repo(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    root = tmp_path / "repo"
    for name, text in {
        ".claude/skills/process-meeting/SKILL.md": "# process-meeting\n",
        "docs/changeset.md": "# The edit engine\n",
        "deckcheck/pyproject.toml": "[project]\n",
        "CLAUDE.md": "# pptx\n",
        "private/m1/transcript.md": "[00:00:01] Ana Ruiz (Partner, Firm): Use the new figures.\n",
        "private/m1/notes.md": "Ana: use the new figures\n",
        "private/m1/data/x.csv": "year,sales\n2025,130\n",
        "private/m1/expected.yaml": "changes: []\n",
        "private/m1/changeset.json": "{}\n",
        "private/m1/build.py": "\n",
        "private/m2/transcript.md": "[00:00:01] Another client (CEO, Other): Our numbers.\n",
        "artifacts/run1/stage/stale.md": "left by an earlier run\n",
    }.items():
        (root / name).parent.mkdir(parents=True, exist_ok=True)
        (root / name).write_text(text)
    build_deck(root / DECK)
    shutil.copy(root / DECK, root / "private/m1/after.pptx")
    monkeypatch.chdir(root)
    monkeypatch.setenv("PATH", str(tmp_path / "bin"))
    return root


def fake_claude(repo: Path, logged_in: bool = True, fixture: Path | None = None) -> Path:
    log = repo.parent / "claude-calls.jsonl"
    script = repo.parent / "bin" / "claude"
    script.parent.mkdir(exist_ok=True)
    script.write_text(
        FAKE.format(
            python=sys.executable,
            logged_in=logged_in,
            log=str(log),
            fixture=str(fixture) if fixture else "",
            root=str(repo),
            forbidden=FORBIDDEN,
        )
    )
    script.chmod(0o755)
    return log


def fixture(repo: Path, changes: list[dict], path: str = "before.pptx") -> Path:
    return changeset(repo / DECK, changes, "fixture.json", path=path)


def calls(log: Path) -> list[dict]:
    return [json.loads(line) for line in log.read_text().splitlines()]


def run(argv: list[str], capsys: pytest.CaptureFixture[str]) -> tuple[int, str, str]:
    code = main(argv)
    out, err = capsys.readouterr()
    return code, out, err


@pytest.mark.parametrize(
    ("meeting", "out"),
    [("private/m1", "artifacts/run1"), ("private/q3 steer co", "private/q3 steer co/run 1")],
    ids=["plain", "spaces"],
)
def test_a_valid_changeset_runs_claude_in_a_staged_folder(repo: Path, capsys: pytest.CaptureFixture[str], meeting: str, out: str) -> None:
    if meeting != "private/m1":
        shutil.copytree(repo / "private/m1", repo / meeting)
    log = fake_claude(repo, fixture=fixture(repo, [change("c1", REVENUE), change("c2", CELL)]))

    assert run(["process", meeting, "--out", out, "--shareable"], capsys) == (
        0,
        f"VALID {out}/changeset.json: 2 changes (1 text-only, 1 structural) on slide 1; 1 ask, 0 flags, 0 held\n"
        "  c1 replace_text slide 1 shape 4 'TextBox 3': '12%' -> '15%' (text-only)\n"
        "  c2 set_cell slide 1 shape 5 'Table 4': '$120m' -> '$130m' (structural)\n",
        "",
    )
    [call] = calls(log)
    assert (call["cwd"], call["argv"], call["files"]) == (f"{out}/stage", ARGV, STAGED)
    assert json.loads((repo / out / "changeset.json").read_text())["source"]["path"] == f"{meeting}/before.pptx"
    assert json.loads((repo / out / "claude.json").read_text()) == {"type": "result", "result": "done"}


def test_the_run_is_granted_no_path_outside_its_staged_folder(repo: Path, capsys: pytest.CaptureFixture[str]) -> None:
    log = fake_claude(repo, fixture=fixture(repo, [change("c1", REVENUE)]))

    assert run(ARGS, capsys)[0] == 0
    assert calls(log)[0]["reads"] == {path: None for path in FORBIDDEN}


def test_an_invalid_changeset_lists_its_problems_and_exits_1(repo: Path, capsys: pytest.CaptureFixture[str]) -> None:
    fake_claude(repo, fixture=fixture(repo, [change("c1", {**REVENUE, "old": "13%"})]))

    assert run(ARGS, capsys) == (
        1,
        "INVALID artifacts/run1/changeset.json: 1 problem\n"
        "  c1 op.old: '13%' is not in shape 4 'TextBox 3'; its text is 'Revenue grew 12% in 2025'\n",
        "",
    )


def test_a_changeset_for_another_deck_is_refused(repo: Path, capsys: pytest.CaptureFixture[str]) -> None:
    fake_claude(repo, fixture=fixture(repo, [change("c1", REVENUE)], path="../../../private/m1/after.pptx"))

    assert run(ARGS, capsys) == (
        1,
        "INVALID artifacts/run1/stage/changeset.json: 1 problem\n"
        "  source.path: '../../../private/m1/after.pptx' is not the deck this meeting edits; name before.pptx\n",
        "",
    )
    assert not (repo / "artifacts/run1/changeset.json").exists()


def test_no_changeset_written_exits_1(repo: Path, capsys: pytest.CaptureFixture[str]) -> None:
    (repo / "artifacts/run1/changeset.json").write_text("{}\n")
    fake_claude(repo)

    assert run(ARGS, capsys) == (1, "", "error: claude wrote no ChangeSet at artifacts/run1/stage/changeset.json\n")
    assert not (repo / "artifacts/run1/changeset.json").exists()


def test_no_claude_on_the_path_exits_3(repo: Path, capsys: pytest.CaptureFixture[str]) -> None:
    assert run(ARGS, capsys) == (3, "", "error: claude CLI not found or not logged in; run `claude` once to log in\n")


def test_a_claude_that_is_not_logged_in_exits_3(repo: Path, capsys: pytest.CaptureFixture[str]) -> None:
    log = fake_claude(repo, logged_in=False)

    assert run(ARGS, capsys) == (3, "", "error: claude CLI not found or not logged in; run `claude` once to log in\n")
    assert not log.exists()


@pytest.mark.parametrize(
    ("args", "message"),
    [
        (
            ["process", "private/m1", "--out", "reports/run1", "--shareable"],
            "--out reports/run1 must sit inside artifacts/ or private/, which git ignores, so client files stay out of git",
        ),
        (
            ["process", "private/m1", "--out", "artifacts/run1"],
            "pass --shareable only if this deck and meeting may be shared with an AI service under your firm's policy "
            "(CLAUDE.md); the run sends them to Claude",
        ),
        (
            ["process", "private/m1/stage", "--out", "private/m1", "--shareable"],
            "private/m1/stage is replaced on every run, so the meeting and the deck cannot sit inside it",
        ),
    ],
    ids=["out-outside-ignored-folders", "not-shareable", "meeting-inside-stage"],
)
def test_a_refused_job_exits_2_and_runs_nothing(repo: Path, capsys: pytest.CaptureFixture[str], args: list[str], message: str) -> None:
    shutil.copytree(repo / "private/m1", repo / "private/m1/stage")
    log = fake_claude(repo)

    assert run(args, capsys) == (2, "", f"error: {message}\n")
    assert not log.exists()
    assert (repo / "private/m1/stage/transcript.md").is_file()
