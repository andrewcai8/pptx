from __future__ import annotations

import json
import shutil
import sys
from pathlib import Path

import pytest
from test_changeset import CELL, REVENUE, build_deck, change, changeset

from deckcheck.meeting import main

DECK = "private/m1/before.pptx"
FAKE = """#!{python}
import json, shutil, sys
from pathlib import Path

args = sys.argv[1:]
if args[:2] == ["auth", "status"]:
    print(json.dumps({{"loggedIn": {logged_in}, "authMethod": "none"}}))
    sys.exit(0 if {logged_in} else 1)
with open({log!r}, "a") as log:
    log.write(json.dumps(args) + "\\n")
prefix = "Bash(uv run --project deckcheck changeset validate "
out = next(a[len(prefix):-1] for a in args if a.startswith(prefix))
if {fixture!r}:
    shutil.copy({fixture!r}, out)
print(json.dumps({{"type": "result", "result": "done"}}))
"""
PROMPT = """Follow the skill at .claude/skills/process-meeting/SKILL.md to turn this meeting into a ChangeSet for its deck.

- Meeting folder: private/m1
- Transcript: private/m1/transcript.md
- Data files: private/m1/data/x.csv
- Source deck: private/m1/before.pptx
- Working folder: artifacts/run1
- ChangeSet: write it to artifacts/run1/changeset.json

These are the only commands you can run. Run each exactly as written:
- `uv run --project deckcheck changeset outline private/m1/before.pptx --json artifacts/run1/outline.json`
- `uv run --project deckcheck changeset validate artifacts/run1/changeset.json`
- `uv run --project deckcheck changeset execute artifacts/run1/changeset.json --out artifacts/run1/executed.pptx --review artifacts/run1/review.json`
- `uv run --project deckcheck deckcheck check artifacts/run1/executed.pptx --out artifacts/run1/check`

The user ran `meeting process --shareable`, which says this deck and meeting may be shared with an AI service under their firm's policy. Do not ask again.

Nobody can answer questions during this run. Put anything you cannot settle in the ChangeSet's flags.
"""
ARGS = ["process", "private/m1", "--out", "artifacts/run1", "--shareable"]


@pytest.fixture
def repo(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    root = tmp_path / "repo"
    skill = root / ".claude/skills/process-meeting/SKILL.md"
    skill.parent.mkdir(parents=True)
    skill.write_text("# process-meeting\n")
    (root / "private/m1/data").mkdir(parents=True)
    (root / "private/m1/transcript.md").write_text("00:00:01 Ana Ruiz: Use the new figures.\n")
    (root / "private/m1/data/x.csv").write_text("year,sales\n2025,130\n")
    build_deck(root / DECK)
    monkeypatch.chdir(root)
    monkeypatch.setenv("PATH", str(tmp_path / "bin"))
    return root


def fake_claude(repo: Path, logged_in: bool = True, fixture: Path | None = None) -> Path:
    log = repo.parent / "claude-calls.jsonl"
    script = repo.parent / "bin" / "claude"
    script.parent.mkdir(exist_ok=True)
    script.write_text(FAKE.format(python=sys.executable, logged_in=logged_in, log=str(log), fixture=str(fixture) if fixture else ""))
    script.chmod(0o755)
    return log


def fixture(repo: Path, changes: list[dict], path: str = DECK) -> Path:
    return changeset(repo / DECK, changes, "fixture.json", path=path)


def run(argv: list[str], capsys: pytest.CaptureFixture[str]) -> tuple[int, str, str]:
    code = main(argv)
    out, err = capsys.readouterr()
    return code, out, err


def test_a_valid_changeset_runs_claude_with_the_documented_invocation(repo: Path, capsys: pytest.CaptureFixture[str]) -> None:
    log = fake_claude(repo, fixture=fixture(repo, [change("c1", REVENUE), change("c2", CELL)]))

    assert run(ARGS, capsys) == (
        0,
        "VALID artifacts/run1/changeset.json: 2 changes (1 text-only, 1 structural) on slide 1; 1 ask, 0 flags, 0 held\n"
        "  c1 replace_text slide 1 shape 4 'TextBox 3': '12%' -> '15%' (text-only)\n"
        "  c2 set_cell slide 1 shape 5 'Table 4': '$120m' -> '$130m' (structural)\n",
        "",
    )
    assert [json.loads(line) for line in log.read_text().splitlines()] == [
        [
            "-p",
            PROMPT,
            "--tools",
            "Read,Write,Edit,Bash",
            "--permission-mode",
            "dontAsk",
            "--allowedTools",
            "Read(./private/m1/**)",
            "Read(./.claude/skills/process-meeting/**)",
            "Read(./docs/changeset.md)",
            "Read(./artifacts/run1/**)",
            "Write(./artifacts/run1/**)",
            "Edit(./artifacts/run1/**)",
            "Bash(uv run --project deckcheck changeset outline private/m1/before.pptx --json artifacts/run1/outline.json)",
            "Bash(uv run --project deckcheck changeset validate artifacts/run1/changeset.json)",
            "Bash(uv run --project deckcheck changeset execute artifacts/run1/changeset.json --out artifacts/run1/executed.pptx --review artifacts/run1/review.json)",
            "Bash(uv run --project deckcheck deckcheck check artifacts/run1/executed.pptx --out artifacts/run1/check)",
            "--output-format",
            "json",
            "--no-session-persistence",
        ]
    ]
    assert json.loads((repo / "artifacts/run1/claude.json").read_text()) == {"type": "result", "result": "done"}


def test_an_invalid_changeset_lists_its_problems_and_exits_1(repo: Path, capsys: pytest.CaptureFixture[str]) -> None:
    fake_claude(repo, fixture=fixture(repo, [change("c1", {**REVENUE, "old": "13%"})]))

    assert run(ARGS, capsys) == (
        1,
        "INVALID artifacts/run1/changeset.json: 1 problem\n"
        "  c1 op.old: '13%' is not in shape 4 'TextBox 3'; its text is 'Revenue grew 12% in 2025'\n",
        "",
    )


def test_a_changeset_for_another_deck_is_refused(repo: Path, capsys: pytest.CaptureFixture[str]) -> None:
    shutil.copy(repo / DECK, repo / "private/m1/other.pptx")
    fake_claude(repo, fixture=fixture(repo, [change("c1", REVENUE)], path="private/m1/other.pptx"))

    assert run(ARGS, capsys) == (
        1,
        "INVALID artifacts/run1/changeset.json: 1 problem\n"
        "  source.path: 'private/m1/other.pptx' is not the deck this meeting edits; name private/m1/before.pptx\n",
        "",
    )


def test_no_changeset_written_exits_1(repo: Path, capsys: pytest.CaptureFixture[str]) -> None:
    fake_claude(repo)

    assert run(ARGS, capsys) == (1, "", "error: claude wrote no ChangeSet at artifacts/run1/changeset.json\n")


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
    ],
    ids=["out-outside-ignored-folders", "not-shareable"],
)
def test_a_refused_job_exits_2_and_runs_nothing(repo: Path, capsys: pytest.CaptureFixture[str], args: list[str], message: str) -> None:
    log = fake_claude(repo)

    assert run(args, capsys) == (2, "", f"error: {message}\n")
    assert not log.exists()
