"""Prove the golden scenarios and their scorer: build every variant, score it, and check each verdict is the declared one.

Usage, from the repo root:
    uv run --project deckcheck python evals/prove.py [names...] [--out RUN_DIR]

With no names it proves every scenario in evals/ and checks the set's composition. Private scenarios are scored with
score.py, never proved here. Exit 0 on PROOF PASS, 1 on a verdict mismatch, 2 on a bad scenario or build, 3 on an
unreachable deck.
"""

from __future__ import annotations

import argparse
import contextlib
import os
import sys
import time
import zipfile
from dataclasses import dataclass
from pathlib import Path

from deckcheck.rules import ConfigError
from deckedit import Variant, build_all
from scenario import (
    ROOT,
    AddSlide,
    Ambiguous,
    BadScenario,
    DeckUnreachable,
    DeleteSlide,
    FromData,
    MoveSlide,
    NotAChange,
    Scenario,
    SourceTampered,
    open_scenario,
    public,
    sha256,
)
from score import BAD, FAIL, OK, UNREACHABLE, Code, SlideRef, line, load_house_rules, run

WORDS = (600, 1500)
NEGATIVE_CODES = frozenset({Code.SCOPE, Code.NON_CHANGE, Code.GUESSED, Code.MISSING, Code.FORBIDDEN, Code.STRUCTURE, Code.LAYOUT, Code.STYLE})


@dataclass(frozen=True)
class Proven:
    scenario: Scenario
    variants: tuple[Variant, ...]


def pairs_text(pairs: frozenset[tuple[Code, SlideRef]]) -> str:
    return "{" + ", ".join(f"{code} {slide}" for code, slide in sorted(pairs, key=lambda p: (list(Code).index(p[0]), str(p[1])))) + "}"


def zip_entries(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path) as z:
        return {info.filename: z.read(info) for info in z.infolist()}


def prove_scenario(sc: Scenario, run_dir: Path, rules) -> tuple[list[str], tuple[Variant, ...]]:
    problems: list[str] = []
    log = run_dir / sc.name / "build.log"
    log.parent.mkdir(parents=True, exist_ok=True)
    with open(log, "w") as f, contextlib.redirect_stderr(f):
        built = build_all(sc, run_dir)
    variants = tuple(v for v, _ in built)
    for v, path in built:
        verdict = run(sc, path, rules, path.parent)
        got = "PASS" if verdict.passed else f"FAIL {pairs_text(verdict.pairs)}"
        ok = verdict.passed if not v.fails else (not verdict.passed and verdict.pairs == v.fails)
        print(f"{'ok  ' if ok else 'BAD '} {sc.name}/{v.name}  {got}")
        if not ok:
            want = pairs_text(v.fails) if v.fails else "PASS"
            print(f"     declared {want}; {line(sc.name, verdict, len(sc.deferred))}")
            problems.append(f"{sc.name}/{v.name} scored {got}, declared {want}")
    passing = [v for v in variants if not v.fails]
    if not passing:
        problems.append(f"{sc.name} has no passing variant")
    if len(variants) - len(passing) < 2:
        problems.append(f"{sc.name} has {len(variants) - len(passing)} failing variants, expected at least 2")
    paths = {v.name: p for v, p in built}
    references = {"the source deck": zip_entries(sc.source_path)} | {v.name: zip_entries(paths[v.name]) for v in passing}
    for v in variants:
        if v.fails:
            entries = zip_entries(paths[v.name])
            problems += [f"{sc.name}/{v.name} has the same content as {name}" for name, ref in references.items() if entries == ref]
    return problems + own_composition(sc), variants


def own_composition(sc: Scenario) -> list[str]:
    problems = []
    if not WORDS[0] <= sc.words <= WORDS[1]:
        problems.append(f"{sc.name} transcript has {sc.words} words, expected {WORDS[0]} to {WORDS[1]}")
    if not any(isinstance(nc, NotAChange) for nc in sc.non_changes):
        problems.append(f"{sc.name} has no not-a-change")
    if not any(f.superseded for _, _, facts in sc.fact_targets for f in facts.forbid):
        problems.append(f"{sc.name} has no superseded forbid (a change of mind)")
    return problems


def set_composition(proven: list[Proven]) -> list[str]:
    scenarios = [p.scenario for p in proven]
    changes = [c for sc in scenarios for c in sc.changes]
    decks = {sc.ref for sc in scenarios}
    codes = {code for p in proven for v in p.variants for code, _ in v.fails}
    problems = []
    if len(scenarios) < 5:
        problems.append(f"the public set has {len(scenarios)} scenarios, expected at least 5")
    if len(decks) < 4:
        problems.append(f"the public set uses {len(decks)} distinct decks, expected at least 4")
    if not any(isinstance(nc, Ambiguous) for sc in scenarios for nc in sc.non_changes):
        problems.append("no public scenario has an ambiguous non-change")
    if not any(isinstance(c, AddSlide) for c in changes):
        problems.append("no public scenario adds a slide")
    if not any(isinstance(c, DeleteSlide | MoveSlide) for c in changes):
        problems.append("no public scenario deletes or moves a slide")
    if not any(isinstance(r.source, FromData) for sc in scenarios for _, _, facts in sc.fact_targets for r in facts.require):
        problems.append("no public scenario takes a number from a data file")
    if uncovered := sorted(NEGATIVE_CODES - codes):
        problems.append(f"no public negative declares {', '.join(uncovered)}")
    return problems


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="prove.py", description="Prove the golden scenarios and their scorer.")
    parser.add_argument("names", nargs="*", help="public scenario names in evals/ (default: all)")
    parser.add_argument("--out", type=Path, help="run directory (default: artifacts/evals/prove-<time>-<pid>)")
    args = parser.parse_args(argv)
    run_dir = (args.out or ROOT / f"artifacts/evals/prove-{time.strftime('%Y%m%d-%H%M%S')}-{os.getpid()}").resolve()
    status = OK
    problems: list[str] = []
    proven: list[Proven] = []
    try:
        dirs = public(args.names)
        rules = load_house_rules()
    except (BadScenario, ConfigError) as e:
        print(f"PROOF FAIL: {e}")
        print(f"evidence: {run_dir}")
        return BAD
    for d in dirs:
        try:
            sc = open_scenario(d)
            found, variants = prove_scenario(sc, run_dir, rules)
        except SourceTampered as e:
            print(f"BAD  {d.name}  {e}")
            problems.append(f"{d.name} source tampered")
            status = max(status, FAIL)
            continue
        except DeckUnreachable as e:
            print(f"BAD  {d.name}  {e}")
            problems.append(f"{d.name} deck unreachable")
            status = UNREACHABLE
            continue
        except Exception as e:
            print(f"BAD  {d.name}  {type(e).__name__}: {e}")
            problems.append(f"{d.name} is a bad scenario or build")
            status = max(status, BAD)
            continue
        problems += found
        proven.append(Proven(sc, variants))
    for p in proven:
        if (got := sha256(p.scenario.source_path)) != p.scenario.source.deck.sha256:
            problems.append(f"{p.scenario.name} source deck {p.scenario.source_path} now hashes to {got}; a build wrote into it")
    if not args.names:
        problems += set_composition(proven)
    if problems:
        status = max(status, FAIL)
        print("PROOF FAIL: " + "; ".join(problems))
    else:
        n_variants = sum(len(p.variants) for p in proven)
        print(f"PROOF PASS ({len(proven)} scenarios, {n_variants} variants, {len({p.scenario.ref for p in proven})} decks)")
    print(f"evidence: {run_dir}")
    return status


if __name__ == "__main__":
    sys.exit(main())
