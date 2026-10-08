"""Score one output deck against a golden scenario.

Usage, from the repo root:
    uv run --project deckcheck python evals/score.py <scenario> <output.pptx> [--out DIR]

Exit 0 on SCENARIO PASS, 1 on SCENARIO FAIL, 2 on a bad scenario or unreadable input (including SCENARIO UNREADABLE), 3 on an unreachable source deck.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
from collections import Counter
from dataclasses import dataclass, replace
from enum import StrEnum
from functools import cached_property
from pathlib import Path
from typing import Literal

from deckcheck.cli import write_atomic
from deckcheck.diff import diff_decks
from deckcheck.model import DeckError, Slide, Violation
from deckcheck.rules import ConfigError, RuleSet, load_rules, run_rules
from facts import ChartValue, Value, describe, match, to_json
from scenario import (
    AddSlide,
    Ambiguous,
    BadScenario,
    Charts,
    DeckUnreachable,
    FromData,
    FromSaid,
    FromSlide,
    NewSlot,
    NonChange,
    Provenance,
    Scenario,
    Slot,
    Snapshot,
    SourceSlot,
    SourceTampered,
    Where,
    corpus,
    find,
    open_scenario,
    snapshot,
)

OK, FAIL, BAD, UNREACHABLE = 0, 1, 2, 3


class Code(StrEnum):
    SOURCE = "source"
    STRUCTURE = "structure"
    SCOPE = "scope"
    NON_CHANGE = "non-change"
    GUESSED = "guessed"
    MISSING = "missing"
    FORBIDDEN = "forbidden"
    LAYOUT = "layout"
    STYLE = "style"
    UNREADABLE = "unreadable"


SlideRef = int | str | None


@dataclass(frozen=True)
class Failure:
    code: Code
    slide: SlideRef
    ref: str | None
    message: str


@dataclass(frozen=True)
class Changed:
    source: int
    output: int
    allowed_by: tuple[str, ...]
    diff: str


@dataclass(frozen=True)
class FactResult:
    ref: str
    slide: SlideRef
    kind: Literal["require", "forbid", "absent"]
    value: Value
    where: Where
    found: str | None
    ok: bool
    source: Provenance | None = None
    superseded: str | None = None


@dataclass(frozen=True)
class Flag:
    question: str
    said: tuple[str, ...]
    slides: tuple[int, ...]

    def raises(self, nc: Ambiguous) -> bool:
        return bool(set(self.said) & set(nc.said) or set(self.slides) & set(nc.slides))


@dataclass(frozen=True)
class Flags:
    """What the maker's flags.json raised, reported beside the verdict and never part of it."""

    file: Path | None
    raised: tuple[str, ...]
    missing: tuple[str, ...]
    unmatched: tuple[str, ...]
    unreadable: tuple[str, ...]

    @property
    def line(self) -> str:
        found = f"raised {', '.join(self.raised) or 'none'}; missing {', '.join(self.missing) or 'none'}; {len(self.unmatched)} unmatched"
        if self.unreadable:
            found += f"; unreadable ({'; '.join(self.unreadable)})"
        return f"flags (reported, not scored): {found}" + ("" if self.file else "; no flags.json")


@dataclass(frozen=True)
class Placement:
    slots: tuple[Slot | None, ...]

    @cached_property
    def index(self) -> dict[Slot, int]:
        return {slot: i for i, slot in enumerate(self.slots) if slot is not None}


@dataclass(frozen=True)
class Verdict:
    failures: tuple[Failure, ...]
    checks: int
    placement: Placement | None = None
    changed: tuple[Changed, ...] = ()
    facts: tuple[FactResult, ...] = ()
    style_inherited: int = 0
    style_new: tuple[tuple[SlideRef, Violation], ...] = ()
    flags: Flags | None = None

    @property
    def passed(self) -> bool:
        return not self.failures

    @property
    def status(self) -> Literal["PASS", "FAIL", "UNREADABLE"]:
        """A deck that fails any decided check fails. Otherwise a check the scorer could not decide leaves it unreadable."""
        if not self.failures:
            return "PASS"
        return "UNREADABLE" if all(f.code == Code.UNREADABLE for f in self.failures) else "FAIL"

    @property
    def pairs(self) -> frozenset[tuple[Code, SlideRef]]:
        return frozenset((f.code, f.slide) for f in self.failures)


def slot_ref(slot: Slot | None) -> SlideRef:
    if slot is None:
        return None
    return slot.slide if isinstance(slot, SourceSlot) else slot.change


def slide_name(ref: SlideRef) -> str:
    if ref is None:
        return "an extra slide"
    return f"slide {ref}" if isinstance(ref, int) else f"the slide {ref} adds"


def failure_key(f: Failure) -> tuple:
    slide = (0, f.slide, "") if isinstance(f.slide, int) else (1, 0, f.slide or "")
    return (list(Code).index(f.code), slide, f.ref or "", f.message)


def sorted_failures(failures: list[Failure]) -> tuple[Failure, ...]:
    return tuple(sorted(set(failures), key=failure_key))


def line(sc_name: str, verdict: Verdict, deferred: int) -> str:
    if verdict.passed:
        return f"SCENARIO PASS {sc_name} ({verdict.checks} checks, {deferred} intent checks deferred)"
    return f"SCENARIO {verdict.status}: " + "; ".join(f"[{f.code}] {f.message}" for f in verdict.failures)


def place_name(ref: SlideRef, where: Where) -> str:
    return f"the title of {slide_name(ref)}" if where == "title" else slide_name(ref)


def non_change_code(nc: NonChange) -> Code:
    return Code.GUESSED if isinstance(nc, Ambiguous) else Code.NON_CHANGE


def non_change_reason(nc: NonChange) -> str:
    return f"{nc.id} is ambiguous (flag it, do not guess)" if isinstance(nc, Ambiguous) else f"{nc.id} is not a deck change"


def score(sc: Scenario, out: Snapshot, rules: RuleSet) -> Verdict:
    if not set(out.ids) & set(sc.source.ids):
        message = "the output does not keep the source deck's slide ids; edit a copy of the source deck"
        return Verdict((Failure(Code.STRUCTURE, None, None, message),), checks=1)
    placement = place(sc, out)
    structure = check_structure(sc, placement)
    scope, changed = check_scope(sc, out, placement, rules.params["source-on-data-slides"]["prefix"])
    applied = {c.source for c in changed if c.diff or any(e.kind == "restyle" for e in sc.edits.get(c.source, ()))}
    facts, fact_failures = check_facts(sc, out, placement, applied)
    layout = check_layout(sc, out, placement)
    inherited, new = check_style(sc, out, placement, rules)
    style = [
        Failure(Code.STYLE, ref, ref if isinstance(ref, str) else None, f"{slide_name(ref)} [{v.rule}] {v.message}" + (f" | {v.evidence}" if v.evidence else ""))
        for ref, v in new
    ]
    checks = 3 + len(sc.skeleton) + len(facts)
    failures = sorted_failures(structure + scope + fact_failures + layout + style)
    return Verdict(failures, checks, placement, tuple(changed), tuple(facts), inherited, tuple(new))


def place(sc: Scenario, out: Snapshot) -> Placement:
    source = {slide_id: k for k, slide_id in enumerate(sc.source.ids, start=1)}
    added = iter(s for s in sc.skeleton if isinstance(s, NewSlot))
    return Placement(tuple(SourceSlot(source[i]) if i in source else next(added, None) for i in out.ids))


def check_structure(sc: Scenario, placement: Placement) -> list[Failure]:
    at = placement.index
    failures: list[Failure] = []
    for slot in sc.skeleton:
        if slot not in at:
            failures.append(Failure(Code.STRUCTURE, slot_ref(slot), slot_ref(slot) if isinstance(slot, NewSlot) else None, f"{slide_name(slot_ref(slot))} is missing from the output (expected at output {sc.position[slot]})"))
    for k, c in sc.deleted.items():
        if SourceSlot(k) in at:
            failures.append(Failure(Code.STRUCTURE, k, c.id, f"slide {k} is still in the output at position {at[SourceSlot(k)] + 1}, but {c.id} deletes it"))
    for p, slot in enumerate(placement.slots, start=1):
        if slot is None:
            failures.append(Failure(Code.STRUCTURE, None, None, f"output {p} is an extra slide that no change asks for"))
    expected = [s for s in sc.skeleton if s in at]
    actual = [s for s in placement.slots if s in sc.position]
    for slot in out_of_order(expected, actual, lambda s: 1 if isinstance(s, NewSlot) or s.slide in sc.moved else 2):
        match slot:
            case SourceSlot(k) if k in sc.moved:
                how = f"{sc.moved[k].id} moves it after slide {sc.moved[k].after}"
            case SourceSlot(k):
                how = "no change moves it"
            case NewSlot(c):
                how = f"{c} adds it after slide {next(a.after for a in sc.changes if isinstance(a, AddSlide) and a.id == c)}"
        ref = slot_ref(slot)
        failures.append(Failure(Code.STRUCTURE, ref, ref if isinstance(ref, str) else None, f"{slide_name(ref)} is at output {at[slot] + 1}, expected output {sc.position[slot]} ({how})"))
    for f in list(failures):
        if isinstance(f.slide, int) and f.slide in sc.named and f.slide not in sc.moved:
            nc = sc.named[f.slide]
            failures.append(Failure(non_change_code(nc), f.slide, nc.id, f"slide {f.slide} was deleted or moved, but {non_change_reason(nc)}"))
    return failures


# A weighted LCS: when two slides could explain one displacement, a slide a change placed (moved or added) weighs
# less, so the alignment keeps the other slides in place and blames the placed one.
def out_of_order(expected: list[Slot], actual: list[Slot], weight) -> list[Slot]:
    n, m = len(expected), len(actual)
    best = [[0] * (m + 1) for _ in range(n + 1)]
    for i in range(n - 1, -1, -1):
        for j in range(m - 1, -1, -1):
            match = best[i + 1][j + 1] + weight(expected[i]) if expected[i] == actual[j] else -1
            best[i][j] = max(match, best[i + 1][j], best[i][j + 1])
    kept, i, j = set(), 0, 0
    while i < n and j < m:
        if expected[i] == actual[j] and best[i][j] == best[i + 1][j + 1] + weight(expected[i]):
            kept.add(expected[i])
            i, j = i + 1, j + 1
        elif best[i + 1][j] >= best[i][j + 1]:
            i += 1
        else:
            j += 1
    return [s for s in expected if s not in kept]


def check_scope(sc: Scenario, out: Snapshot, placement: Placement, source_prefix: str) -> tuple[list[Failure], list[Changed]]:
    pairs = [(s.slide, placement.index[s]) for s in sc.skeleton if isinstance(s, SourceSlot) and s in placement.index]
    old = replace(sc.source.deck, slides=tuple(sc.source.deck.slides[k - 1] for k, _ in pairs))
    new = replace(out.deck, slides=tuple(out.deck.slides[i] for _, i in pairs))
    failures: list[Failure] = []
    changed: list[Changed] = []
    for (k, i), d in zip(pairs, diff_decks(old, new).slides, strict=True):
        before, after = sc.source.looks[k - 1], out.looks[i]
        text_changed = d.status == "changed"
        look_changed = text_changed or before != after
        edits = sc.edits.get(k, ())
        if look_changed:
            changed.append(Changed(k, i + 1, tuple(e.id for e in edits), d.diff))
        for e in edits:
            if e.kind == "restyle" and text_changed:
                failures.append(Failure(Code.SCOPE, k, e.id, f"{e.id} restyles slide {k} but its text changed"))
            elif e.kind == "restyle" and not look_changed:
                failures.append(Failure(Code.MISSING, k, e.id, f"{e.id} not applied to slide {k} (no restyle)"))
            elif e.kind != "restyle" and not text_changed:
                failures.append(Failure(Code.MISSING, k, e.id, f"{e.id} not applied to slide {k}"))
        if edits and text_changed:
            forbids = [f.value for e in edits for f in e.slides[k].forbid if f.superseded is None and not isinstance(f.value, ChartValue)]
            growth = {g.shape: g.adds for e in edits for g in e.slides[k].may_change}
            lines = unasked_lines(sc.source.deck.slides[k - 1], out.deck.slides[i], forbids, source_prefix, growth)
            by = edits[0].id
            if lines.lost:
                failures.append(Failure(Code.SCOPE, k, by, f"{by} does not ask to change these lines on slide {k}, but they are gone or changed: {shown(lines.lost)}"))
            if lines.added:
                failures.append(Failure(Code.SCOPE, k, by, f"{by} does not ask to add these lines to slide {k}: {shown(lines.added)}"))
            for shape, grown in lines.off_topic.items():
                failures.append(Failure(Code.SCOPE, k, by, f"{by} lets {shape} on slide {k} gain only text that states {describe(growth[shape])}, but it gains: {shown(grown)}"))
        if edits or not look_changed:
            continue
        what = "changed its text" if text_changed else "changed its XML with the same text" if before.xml != after.xml else "changed a chart, image, media, or notes part"
        nc = sc.named.get(k)
        why = f" ({nc.id}: {nc.why})" if nc else ""
        failures.append(Failure(non_change_code(nc) if nc else Code.SCOPE, k, nc.id if nc else None, f"slide {k} {what}, but no change asks for it{why}"))
    if sc.source.shared != out.shared:
        failures.append(Failure(Code.SCOPE, None, None, "the deck's slide layouts, masters, or themes changed, but no change asks for it"))
    return failures, changed


def shown(lines: list[str]) -> str:
    return ", ".join(repr(x) for x in lines[:3]) + (f", and {len(lines) - 3} more" if len(lines) > 3 else "")


@dataclass(frozen=True)
class Unasked:
    lost: list[str]
    added: list[str]
    off_topic: dict[str, list[str]]


def unasked_lines(before: Slide, after: Slide, forbids: list[Value], source_prefix: str, growth: dict[str, Value]) -> Unasked:
    """The lines an edited slide changed beyond its edit, compared shape by shape.

    A source line must survive unless it is released (it holds an old value or is a source line), and a released line
    may be replaced by one new line in its shape. In a shape that may grow, every source line must survive, extended or
    not, and every added stretch of text must state what the growth names. A blank paragraph adds no text.
    """
    def released(text: str) -> bool:
        return text.startswith(source_prefix) or any(match(f, text) for f in forbids)

    old, new = _lines(before), _lines(after)
    lost: list[str] = []
    added: list[str] = []
    off_topic: dict[str, list[str]] = {}
    for name in [*old, *(n for n in new if n not in old)]:
        fresh = list(new.get(name, []))
        gone = []
        for text in old.get(name, []):
            if text in fresh:
                fresh.remove(text)
            else:
                gone.append(text)
        fresh = [text for text in fresh if text]
        if name in growth:
            off = []
            for text in gone:
                if grown := next((f for f in fresh if text and text in f), None):
                    fresh.remove(grown)
                    if any(any(ch.isalnum() for ch in piece) and not match(growth[name], piece) for piece in grown.split(text, 1)):
                        off.append(grown)
                elif not released(text):
                    lost.append(f"{name}: {text}")
            off += [f for f in fresh if not match(growth[name], f)]
            if off:
                off_topic[name] = off
            continue
        lost += [f"{name}: {text}" for text in gone if not released(text)]
        added += [f"{name}: {text}" for text in fresh[sum(map(released, gone)) :]]
    return Unasked(lost, added, off_topic)


def _lines(slide: Slide) -> dict[str, list[str]]:
    lines: dict[str, list[str]] = {}
    for shape in slide.shapes:
        lines.setdefault(shape.name, []).extend(p.text.strip() for p in shape.paragraphs)
    return lines


def check_facts(sc: Scenario, out: Snapshot, placement: Placement, applied: set[int]) -> tuple[list[FactResult], list[Failure]]:
    results: list[FactResult] = []
    failures: list[Failure] = []

    def unreadable(value: Value, charts: Charts, ref: SlideRef, by: str) -> bool:
        if isinstance(value, ChartValue) and charts.unreadable:
            failures.append(Failure(Code.UNREADABLE, ref, by, f"{slide_name(ref)} chart: {', '.join(charts.unreadable)} cannot be read; check by hand"))
            return True
        return False

    for change, slot, facts in sc.fact_targets:
        if slot not in placement.index or isinstance(slot, SourceSlot) and slot.slide not in applied:
            continue
        i = placement.index[slot]
        slide, charts = out.deck.slides[i], out.charts[i]
        ref = slot_ref(slot)
        for r in facts.require:
            if unreadable(r.value, charts, ref, change.id):
                continue
            found = find(r.value, r.where, slide, charts)
            results.append(FactResult(change.id, ref, "require", r.value, r.where, found, found is not None, source=r.source))
            if found is None:
                failures.append(Failure(Code.MISSING, ref, change.id, f"{change.id}: {describe(r.value)} is not on {place_name(ref, r.where)}"))
        for f in facts.forbid:
            if unreadable(f.value, charts, ref, change.id):
                continue
            found = find(f.value, f.where, slide, charts)
            results.append(FactResult(change.id, ref, "forbid", f.value, f.where, found, found is None, superseded=f.superseded))
            if found and f.superseded:
                failures.append(Failure(Code.FORBIDDEN, ref, change.id, f"{change.id}: {found!r} is on {place_name(ref, f.where)}; it was abandoned after {f.superseded}"))
            elif found:
                failures.append(Failure(Code.FORBIDDEN, ref, change.id, f"{change.id}: {found!r} is still on {place_name(ref, f.where)}"))
    for nc in sc.non_changes:
        for value in nc.absent:
            hits = [
                (slot_ref(slot), found)
                for slot, s, charts in zip(placement.slots, out.deck.slides, out.charts, strict=True)
                if not unreadable(value, charts, slot_ref(slot), nc.id) and (found := find(value, "slide", s, charts))
            ]
            results.append(FactResult(nc.id, "deck", "absent", value, "slide", hits[0][1] if hits else None, not hits))
            failures += [Failure(non_change_code(nc), ref, nc.id, f"{nc.id}: {found!r} is on {slide_name(ref)}, but {non_change_reason(nc)}") for ref, found in hits]
    return results, failures


def check_layout(sc: Scenario, out: Snapshot, placement: Placement) -> list[Failure]:
    failures = []
    for c in sc.changes:
        if isinstance(c, AddSlide) and NewSlot(c.id) in placement.index:
            got = out.deck.slides[placement.index[NewSlot(c.id)]].layout_name
            if got != c.layout:
                failures.append(Failure(Code.LAYOUT, c.id, c.id, f"the slide {c.id} adds uses layout {got!r}, expected {c.layout!r}"))
    return failures


def check_style(sc: Scenario, out: Snapshot, placement: Placement, rules: RuleSet) -> tuple[int, list[tuple[SlideRef, Violation]]]:
    baseline = Counter((v.slide, v.rule, v.shape, v.evidence) for v in run_rules(sc.source.deck, rules))
    inherited = 0
    new: list[tuple[SlideRef, Violation]] = []
    for v in run_rules(out.deck, rules):
        ref = slot_ref(placement.slots[v.slide - 1])
        key = (ref, v.rule, v.shape, v.evidence)
        if baseline[key] > 0:
            baseline[key] -= 1
            inherited += 1
        else:
            new.append((ref, v))
    return inherited, new


FLAGS_FILE = "flags.json"
# A turn as a maker may write it, with or without the hour: 00:08:05, 0:08:05, 08:05, or 8:05.
TURN_STAMP = re.compile(r"(?:(\d{1,2}):)?([0-5]?\d):([0-5]\d)")
SLIDE_NUMBER = re.compile(r"\d+", re.ASCII)


class UnreadableFlag(Exception):
    pass


def read_flags(path: Path) -> tuple[tuple[Flag, ...], tuple[str, ...]] | None:
    """The maker's flags.json, if it wrote one: the flags it could read and why it could not read the rest. Flags are
    reported beside the verdict, so a file the scorer cannot read never stops it scoring the deck."""
    if not path.is_file():
        return None
    try:
        raw = json.loads(path.read_text(encoding="utf-8-sig"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as e:
        return (), (f"{path.name}: {e}",)
    if not isinstance(raw, list):
        return (), (f"{path.name}: expected a list of flags",)
    flags: list[Flag] = []
    unreadable: list[str] = []
    for i, item in enumerate(raw):
        try:
            flags.append(parse_flag(item, f"{path.name}[{i}]"))
        except UnreadableFlag as e:
            unreadable.append(str(e))
    return tuple(flags), tuple(unreadable)


def parse_flag(raw: object, where: str) -> Flag:
    """A flag needs its question and what it is about. Any other key is the maker's own and is ignored."""
    if not isinstance(raw, dict):
        raise UnreadableFlag(f"{where}: expected an object with a question, and said or slides")
    question = raw.get("question")
    if not isinstance(question, str) or not question.strip():
        raise UnreadableFlag(f"{where}.question: expected the question as a non-empty string")
    said = tuple(_turn(at, f"{where}.said") for at in _items(raw.get("said")))
    slides = tuple(_slide_number(k, f"{where}.slides") for k in _items(raw.get("slides")))
    if not said and not slides:
        raise UnreadableFlag(f"{where}: name the transcript turns in said or the source slides in slides")
    return Flag(question, said, slides)


def _items(raw: object) -> list[object]:
    return [] if raw is None else raw if isinstance(raw, list) else [raw]


def _turn(raw: object, where: str) -> str:
    m = TURN_STAMP.fullmatch(raw.strip()) if isinstance(raw, str) else None
    if m is None:
        raise UnreadableFlag(f"{where}: {raw!r} is not a transcript timestamp such as 00:08:05")
    hours, minutes, seconds = m.groups()
    return f"{int(hours or 0):02}:{int(minutes):02}:{seconds}"


def _slide_number(raw: object, where: str) -> int:
    if type(raw) is int:
        return raw
    if isinstance(raw, str) and SLIDE_NUMBER.fullmatch(raw.strip()):
        return int(raw)
    raise UnreadableFlag(f"{where}: {raw!r} is not a source slide number")


def check_flags(sc: Scenario, path: Path) -> Flags:
    read = read_flags(path)
    flags, unreadable = read or ((), ())
    asks = [nc for nc in sc.non_changes if isinstance(nc, Ambiguous)]
    raised = tuple(nc.id for nc in asks if any(f.raises(nc) for f in flags))
    return Flags(
        path if read is not None else None,
        raised,
        tuple(nc.id for nc in asks if nc.id not in raised),
        tuple(f.question for f in flags if not any(f.raises(nc) for nc in asks)),
        unreadable,
    )


def provenance_json(p: Provenance | None) -> dict[str, object] | None:
    match p:
        case FromSaid(at):
            return {"said": at}
        case FromData(data, row, column):
            return {"data": data, "row": row, "column": column}
        case FromSlide(k):
            return {"slide": k}
    return None


def failure_json(f: Failure) -> dict[str, object]:
    return {"code": str(f.code), "slide": f.slide, "ref": f.ref, "message": f.message}


def report(sc: Scenario, output: Path, out: Snapshot | None, verdict: Verdict) -> dict[str, object]:
    return {
        "scenario": sc.name,
        "verdict": verdict.status,
        "line": line(sc.name, verdict, len(sc.deferred)),
        "source": {"ref": sc.ref, "path": str(sc.source_path), "sha256": sc.source.deck.sha256, "untouched": True},
        "output": {"path": str(output), "sha256": out.deck.sha256 if out else None, "slides": len(out.deck.slides) if out else None},
        "skeleton": {str(i): slot_ref(s) for i, s in enumerate(sc.skeleton, start=1)},
        "placed": {str(i): slot_ref(s) for i, s in enumerate(verdict.placement.slots, start=1)} if verdict.placement else None,
        "must_stay_unchanged": sorted(sc.frozen),
        "changed": [{"source": c.source, "output": c.output, "allowed_by": list(c.allowed_by) or None, "diff": c.diff} for c in verdict.changed],
        "facts": [
            {
                "ref": f.ref,
                "slide": f.slide,
                "kind": f.kind,
                "fact": to_json(f.value),
                "where": f.where,
                "found": f.found,
                "ok": f.ok,
                "from": provenance_json(f.source),
                "superseded": f.superseded,
            }
            for f in verdict.facts
        ],
        "house_style": {
            "inherited": verdict.style_inherited,
            "new": [{"slide": ref, "rule": v.rule, "shape": v.shape, "message": v.message, "evidence": v.evidence} for ref, v in verdict.style_new],
        },
        "failures": [failure_json(f) for f in verdict.failures],
        "deferred": [{"id": i, "check": text} for i, text in sc.deferred],
        "flags": flags_json(verdict.flags),
        "checks": verdict.checks,
    }


def flags_json(flags: Flags | None) -> dict[str, object] | None:
    if flags is None:
        return None
    return {
        "file": str(flags.file) if flags.file else None,
        "raised": list(flags.raised),
        "missing": list(flags.missing),
        "unmatched": list(flags.unmatched),
        "unreadable": list(flags.unreadable),
    }


def tampered_report(e: SourceTampered, output: Path, verdict: Verdict) -> dict[str, object]:
    return {
        "scenario": e.name,
        "verdict": "FAIL",
        "line": line(e.name, verdict, 0),
        "source": {"ref": e.ref, "path": str(e.path), "sha256": e.got, "untouched": False},
        "output": {"path": str(output), "sha256": None, "slides": None},
        "failures": [failure_json(f) for f in verdict.failures],
        "checks": verdict.checks,
    }


def write_json(path: Path, doc: dict[str, object]) -> None:
    write_atomic(path, (json.dumps(doc, indent=2, ensure_ascii=False) + "\n").encode())


def run(sc: Scenario, output: Path, rules: RuleSet, out_dir: Path) -> Verdict:
    if output.exists() and os.path.samefile(output, sc.source_path):
        verdict = Verdict((Failure(Code.SOURCE, None, None, f"the output {output} is the source deck itself"),), checks=1)
        write_json(out_dir / "score.json", report(sc, output, None, verdict))
        return verdict
    flags = check_flags(sc, output.parent / FLAGS_FILE)
    out = snapshot(output)
    verdict = replace(score(sc, out, rules), flags=flags)
    write_json(out_dir / "score.json", report(sc, output, out, verdict))
    return verdict


def load_house_rules() -> RuleSet:
    return load_rules(corpus.RULES)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="score.py", description="Score an output deck against a golden scenario.")
    parser.add_argument("scenario", help="scenario directory, or a name in evals/ or $GOLDEN_PRIVATE_DIR")
    parser.add_argument("output", type=Path, help="the output .pptx")
    parser.add_argument("--out", type=Path, help="where score.json goes (default: the output's directory)")
    args = parser.parse_args(argv)
    out_dir = args.out or args.output.resolve().parent
    try:
        sc = open_scenario(args.scenario)
        rules = load_house_rules()
        verdict = run(sc, args.output, rules, out_dir)
    except SourceTampered as e:
        verdict = Verdict((Failure(Code.SOURCE, None, None, str(e)),), checks=1)
        write_json(out_dir / "score.json", tampered_report(e, args.output, verdict))
        print(line(e.name, verdict, 0))
        return FAIL
    except (BadScenario, ConfigError, DeckError) as e:
        print(f"error: {e}", file=sys.stderr)
        return BAD
    except DeckUnreachable as e:
        print(f"error: {e}", file=sys.stderr)
        return UNREACHABLE
    print(line(sc.name, verdict, len(sc.deferred)))
    if verdict.flags and (verdict.flags.file or verdict.flags.missing):
        print(verdict.flags.line)
    return {"PASS": OK, "FAIL": FAIL, "UNREADABLE": BAD}[verdict.status]


if __name__ == "__main__":
    sys.exit(main())
