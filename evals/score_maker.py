"""Score a maker's ChangeSets against the public golden scenarios.

Usage, from the repo root:
    uv run --project deckcheck python evals/score_maker.py <run-dir> [scenario ...]

For each scenario, reads <run-dir>/<scenario>/changeset.json, executes it to executed.pptx, writes its flags to
flags.json beside the deck, and scores the deck with score.py. It also reports, without scoring, every ref that does
not quote its transcript turn. With no names it scores every public scenario, so one the maker never wrote fails.

Exit 0 when every scenario passes, 1 when any fails, 2 on bad arguments.
"""

from __future__ import annotations

import argparse
import contextlib
import io
import json
import re
import sys
from pathlib import Path

import score
from deckcheck.changeset import slides
from deckcheck.changeset.engine import Checked, Invalid, execute, load
from deckcheck.changeset.model import ChangeSet, Ref
from deckcheck.cli import write_atomic
from deckcheck.model import DeckError
from deckcheck.package import Package
from scenario import BadScenario, public

OK, FAIL, BAD = 0, 1, 2
OUTPUTS = ("executed.pptx", score.FLAGS_FILE, "score.json")
# scenario.TURN keeps the role in the speaker; a ref names the speaker alone.
TURN = re.compile(r"^\[(\d\d:\d\d:\d\d)\] ([^(]+?) \([^)]*\): (.*)$")


def flags_json(checked: Checked) -> list[dict[str, object]]:
    index = {s.id: s.index for s in slides.read_deck(Package(checked.source)).slides}
    return [
        {"question": f.question, "said": [r.t for r in f.refs], "slides": [index[s] for s in f.slides]}
        for f in checked.changeset.flags
    ]


def ref_problem(ref: Ref, turns: dict[str, tuple[str, str]]) -> str | None:
    if ref.t not in turns:
        return "no turn at this time"
    speaker, text = turns[ref.t]
    if speaker != ref.speaker:
        return f"{speaker} speaks this turn, not {ref.speaker}"
    if ref.quote not in text:
        return f"{ref.quote!r} is not in the turn"
    return None


def refs_line(cs: ChangeSet, transcript: Path) -> str:
    turns = {m[1]: (m[2], m[3]) for line in transcript.read_text().splitlines() if (m := TURN.match(line))}
    refs = [(item.id, r) for item in (*cs.asks, *cs.changes, *cs.flags, *cs.held) for r in item.refs]
    bad = [f"{iid} {r.t}: {problem}" for iid, r in refs if (problem := ref_problem(r, turns))]
    return f"refs (reported, not scored): {len(refs) - len(bad)} of {len(refs)} quote their turn" + "".join(f"; {b}" for b in bad)


def score_one(scenario: Path, run_dir: Path) -> tuple[bool, list[str]]:
    out_dir = run_dir / scenario.name
    path = out_dir / "changeset.json"
    for name in OUTPUTS:
        (out_dir / name).unlink(missing_ok=True)
    if not path.is_file():
        return False, [f"CHANGESET MISSING {path}"]
    try:
        checked = load(path)
    except Invalid as e:
        return False, ["CHANGESET INVALID: " + "; ".join(f"{p.where}: {p.message}" for p in e.problems)]
    except DeckError as e:
        return False, [f"CHANGESET UNREADABLE: {e}"]
    executed = out_dir / "executed.pptx"
    write_atomic(executed, execute(checked))
    write_atomic(out_dir / score.FLAGS_FILE, (json.dumps(flags_json(checked), indent=2, ensure_ascii=False) + "\n").encode())
    captured = io.StringIO()
    with contextlib.redirect_stdout(captured), contextlib.redirect_stderr(captured):
        code = score.main([scenario.name, str(executed), "--out", str(out_dir)])
    return code == score.OK, [*captured.getvalue().splitlines(), refs_line(checked.changeset, scenario / "transcript.md")]


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="score_maker.py", description="Score a maker's ChangeSets against the public golden scenarios.")
    parser.add_argument("run_dir", type=Path, help="holds <scenario>/changeset.json for each scenario")
    parser.add_argument("scenarios", nargs="*", help="public scenario names (default: all of them)")
    args = parser.parse_args(argv)
    if not args.run_dir.is_dir():
        print(f"error: no run directory {args.run_dir}", file=sys.stderr)
        return BAD
    try:
        scenarios = public(args.scenarios)
    except BadScenario as e:
        print(f"error: {e}", file=sys.stderr)
        return BAD
    passed = 0
    for scenario in scenarios:
        ok, (first, *rest) = score_one(scenario, args.run_dir)
        passed += ok
        print(f"{scenario.name}: {first}")
        for extra in rest:
            print(f"  {extra}")
    print(f"MAKER {passed} of {len(scenarios)} SCENARIO PASS")
    return OK if passed == len(scenarios) else FAIL


if __name__ == "__main__":
    sys.exit(main())
