from __future__ import annotations

import argparse
import json
import os
import secrets
import sys
from dataclasses import asdict, fields, replace
from importlib.metadata import version
from pathlib import Path

from deckcheck.diff import DeckDiff, diff_decks
from deckcheck.fix import FIXERS, Change, Fixed, FixResult, Reported, fix_deck, plural
from deckcheck.model import Deck, DeckError, Violation, load_deck, read_bytes
from deckcheck.render import RenderError, ToolMissing, find_fc_match, find_pdftoppm, find_soffice, render
from deckcheck.rules import ConfigError, RuleSet, load_rules, run_rules

OK, VIOLATIONS, USAGE, MISSING_TOOL = 0, 1, 2, 3
DEFAULT_RULES = Path("standards/house-style.yaml")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="deckcheck", description="Verify a .pptx deck against the house style.")
    sub = parser.add_subparsers(dest="command", required=True)

    check = sub.add_parser("check", help="check a deck against the house style")
    check.add_argument("deck", type=Path)
    check.add_argument("--rules", type=Path, help=f"rules yaml (default: nearest {DEFAULT_RULES})")
    check.add_argument("--out", type=Path, help="write report.json and outline.md here")

    diff = sub.add_parser("diff", help="compare two decks slide by slide")
    diff.add_argument("old", type=Path)
    diff.add_argument("new", type=Path)
    diff.add_argument("--out", type=Path, help="write diff.json and diff.md here")

    fix = sub.add_parser("fix", help="write a copy of a deck with the fixable violations fixed")
    fix.add_argument("deck", type=Path)
    fix.add_argument("--out", type=Path, required=True, help="the fixed deck, a new .pptx path")
    fix.add_argument("--rules", type=Path, help=f"rules yaml (default: nearest {DEFAULT_RULES})")
    fix.add_argument("--report", type=Path, help="write fix.json and fix.md here")

    rend = sub.add_parser("render", help="render slides to PNG via LibreOffice")
    rend.add_argument("deck", type=Path)
    rend.add_argument("--out", type=Path, required=True)

    sub.add_parser("doctor", help="report versions, rules, and external tools")

    args = parser.parse_args(argv)
    try:
        match args.command:
            case "check":
                return cmd_check(args.deck, args.rules, args.out)
            case "diff":
                return cmd_diff(args.old, args.new, args.out)
            case "fix":
                return cmd_fix(args.deck, args.out, args.rules, args.report)
            case "render":
                return cmd_render(args.deck, args.out)
            case _:
                return cmd_doctor()
    except (ConfigError, DeckError, RenderError) as e:
        print(f"error: {e}", file=sys.stderr)
        return USAGE
    except ToolMissing as e:
        print(f"error: {e}", file=sys.stderr)
        return MISSING_TOOL


def find_rules(start: Path) -> Path:
    for d in (start, *start.parents):
        if (d / DEFAULT_RULES).is_file():
            return d / DEFAULT_RULES
    raise ConfigError(f"no {DEFAULT_RULES} found in {start} or any parent; pass --rules")


def cmd_check(deck_path: Path, rules_path: Path | None, out: Path | None) -> int:
    rules = load_rules(rules_path or find_rules(Path.cwd()))
    deck = load_deck(deck_path)
    violations = run_rules(deck, rules)
    if violations:
        print(f"FAIL {deck_path}: {len(violations)} violations")
        for v in violations:
            print(format_violation(v))
    else:
        print(f"PASS {deck_path} ({len(deck.slides)} slides, {len(rules.params)} rules)")
    if out:
        out.mkdir(parents=True, exist_ok=True)
        report = {
            "deck": str(deck_path),
            "sha256": deck.sha256,
            "rules_path": str(rules.path),
            "rules": list(rules.params),
            "passed": not violations,
            "violations": [violation_json(v) for v in violations],
        }
        (out / "report.json").write_text(json.dumps(report, indent=2) + "\n")
        (out / "outline.md").write_text(outline_md(deck))
    return VIOLATIONS if violations else OK


def violation_json(v: Violation) -> dict[str, object]:
    return {f.name: getattr(v, f.name) for f in fields(v) if f.compare}


def format_violation(v: Violation) -> str:
    where = f" {v.shape}:" if v.shape else ""
    evidence = f" | {v.evidence}" if v.evidence else ""
    return f"slide {v.slide} [{v.rule}]{where} {v.message}{evidence}"


def outline_md(deck: Deck) -> str:
    lines = [f"# {deck.path}", ""]
    for slide in deck.slides:
        lines += [f"## Slide {slide.index}: {slide.title}", "", f"Layout: {slide.layout_name}", ""]
        for shape in slide.shapes:
            if not shape.paragraphs:
                lines.append(f"- {shape.name} ({shape.kind})")
            for p in shape.paragraphs:
                fonts = sorted({f"{r.font} {r.size_pt:g}pt" if r.size_pt else r.font for r in p.runs})
                bullet = "bullet, " if p.is_bullet else ""
                lines.append(f"- {shape.name} ({shape.kind}): {p.text.strip()} [{bullet}{', '.join(fonts)}]")
        lines.append("")
    return "\n".join(lines)


def cmd_diff(old_path: Path, new_path: Path, out: Path | None) -> int:
    result = diff_decks(load_deck(old_path), load_deck(new_path))
    print(f"old {result.old} sha256 {result.old_sha256}")
    print(f"new {result.new} sha256 {result.new_sha256}")
    changes = [s for s in result.slides if s.status != "unchanged"]
    if not changes:
        print("no slide changes")
    for s in changes:
        print(f"slide {s.slide}: {s.status} | {s.title}")
    for s in changes:
        print(s.diff)
    if out:
        out.mkdir(parents=True, exist_ok=True)
        (out / "diff.json").write_text(json.dumps(asdict(result), indent=2) + "\n")
        (out / "diff.md").write_text(diff_md(result))
    return OK


def diff_md(result: DeckDiff) -> str:
    lines = [
        "# Deck diff",
        "",
        f"- old: `{result.old}` sha256 `{result.old_sha256}`",
        f"- new: `{result.new}` sha256 `{result.new_sha256}`",
        "",
    ]
    for s in result.slides:
        lines.append(f"## Slide {s.slide}: {s.status} | {s.title}")
        lines.append("")
        if s.diff:
            lines += ["```diff", s.diff, "```", ""]
    return "\n".join(lines)


def cmd_fix(deck_path: Path, out: Path, rules_path: Path | None, report: Path | None) -> int:
    rules = load_rules(rules_path or find_rules(Path.cwd()))
    data = read_bytes(deck_path)
    if out.is_dir():
        print(f"error: --out {out} is a directory; pass the path of the new deck", file=sys.stderr)
        return USAGE
    if out.exists() and os.path.samefile(out, deck_path):
        print(f"error: --out {out} is the input deck; fix never overwrites its input", file=sys.stderr)
        return USAGE
    try:
        if report:
            report.mkdir(parents=True, exist_ok=True)
        result = fix_deck(data, rules, str(deck_path))
        write_atomic(out, result.data)
        remaining = result.remaining
        print(f"{'FAIL' if remaining else 'PASS'} {deck_path} -> {out}: {fix_summary(result)}")
        for o in (*result.fixed, *remaining):
            print(format_outcome(o))
        if report:
            (report / "fix.json").write_text(json.dumps(fix_json(deck_path, out, rules, result), indent=2) + "\n")
            (report / "fix.md").write_text(fix_md(deck_path, out, rules, result))
    except OSError as e:
        print(f"error: cannot write the fix output: {e}", file=sys.stderr)
        return USAGE
    return VIOLATIONS if remaining else OK


def write_atomic(path: Path, data: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(f".{path.name}.{secrets.token_hex(8)}.part")
    # os.open applies the umask itself; reading it means setting it, which races with other threads.
    fd = os.open(tmp, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o666)
    try:
        with os.fdopen(fd, "wb") as f:
            f.write(data)
        os.replace(tmp, path)
    except BaseException:
        tmp.unlink(missing_ok=True)
        raise


def fix_summary(result: FixResult) -> str:
    passes = f" in {plural(result.passes, 'pass', 'passes')}" if result.passes else ""
    return f"{len(result.fixed)} fixed{passes}, {plural(len(result.remaining), 'remains', 'remain')}"


def format_change(c: Change) -> str:
    if c.what == "text":
        return f"{c.what} {c.before!r} -> {c.after!r}"
    return f"{c.what} {c.before} -> {c.after}"


def format_outcome(o: Fixed | Reported) -> str:
    if isinstance(o, Fixed):
        changes = "; ".join(format_change(c) for c in o.changes)
        return f"fixed {format_violation(replace(o.violation, evidence=changes))}"
    return f"remains {format_violation(o.violation)} ({o.why}: {o.detail})"


def fix_json(deck_path: Path, out: Path, rules: RuleSet, result: FixResult) -> dict[str, object]:
    return {
        "input": str(deck_path),
        "input_sha256": result.input_sha256,
        "output": str(out),
        "output_sha256": result.output_sha256,
        "rules_path": str(rules.path),
        "rules": list(rules.params),
        "passes": result.passes,
        "passed": not result.remaining,
        "fixed": [
            violation_json(o.violation) | {"pass": o.pass_no, "changes": [asdict(c) for c in o.changes]}
            for o in result.fixed
        ],
        "remaining": [violation_json(o.violation) | {"why": o.why, "detail": o.detail} for o in result.remaining],
    }


def fix_md(deck_path: Path, out: Path, rules: RuleSet, result: FixResult) -> str:
    lines = [
        "# Deck fix",
        "",
        f"- input: `{deck_path}` sha256 `{result.input_sha256}`",
        f"- output: `{out}` sha256 `{result.output_sha256}`",
        f"- rules: `{rules.path}` ({', '.join(rules.params)})",
        f"- result: {'PASS' if not result.remaining else 'FAIL'}, {fix_summary(result)}",
        "",
        "## Fixed",
        "",
        *(f"- {format_outcome(o)} (pass {o.pass_no})" for o in result.fixed),
        "",
        "## Remaining",
        "",
        *(f"- {format_outcome(o)}" for o in result.remaining),
        "",
    ]
    return "\n".join(lines)


def cmd_render(deck: Path, out: Path) -> int:
    pngs, fonts = render(deck, out)
    for png in pngs:
        print(png)
    for typeface, match in fonts.items():
        if match.substituted:
            print(f"substituted: {typeface} -> {match.family}")
    return OK


def cmd_doctor() -> int:
    print(f"deckcheck {version('deckcheck')}")
    print(f"python-pptx {version('python-pptx')}")
    status = OK
    try:
        rules = load_rules(find_rules(Path.cwd()))
        print(f"rules: {rules.path}")
        print(f"rule ids: {', '.join(rules.params)}")
    except ConfigError as e:
        print(f"rules: error: {e}")
        status = USAGE
    print(f"fixable rule ids: {', '.join(FIXERS)}")
    print(f"soffice: {find_soffice() or 'missing'}")
    print(f"pdftoppm: {find_pdftoppm() or 'missing'}")
    print(f"fc-match: {find_fc_match() or 'missing'}")
    return status
