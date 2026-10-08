from __future__ import annotations

import json
import re
import sys
from pathlib import Path

PASS_FAIL = ("pass", "fail")
PASS_FAIL_NA = ("pass", "fail", "n/a")
RESULTS = {
    "C1": PASS_FAIL_NA,
    "C2": PASS_FAIL,
    "C3": PASS_FAIL,
    "C4": PASS_FAIL,
    "C5": PASS_FAIL_NA,
    "C6": PASS_FAIL_NA,
    "C7": PASS_FAIL_NA,
    "C8": PASS_FAIL,
}
ADVISORY_WHEN_SUBSTITUTED = ["C3", "C4", "C5"]
VERDICT = {False: "good", True: "needs-work"}
FIELDS = {"slide": int, "checks": dict, "evidence": dict, "advisory": list, "verdict": str}


class Invalid(Exception):
    pass


def load(path: Path) -> object:
    try:
        return json.loads(path.read_text())
    except (OSError, ValueError) as e:
        raise Invalid(f"cannot read {path.name}: {e}") from e


def check(slide_dir: Path, audit_path: Path) -> str:
    if not (slide_dir / "fonts.json").is_file():
        raise Invalid(f"no fonts.json in {slide_dir}")
    fonts = load(slide_dir / "fonts.json")
    if type(fonts) is not dict or any(type(v) is not dict for v in fonts.values()):
        raise Invalid("fonts.json is not an object of font entries")
    substituted = any(v.get("substituted") is True for v in fonts.values())
    want_advisory = ADVISORY_WHEN_SUBSTITUTED if substituted else []
    why = "shows a substituted font" if substituted else "shows no substituted font"
    pngs = sorted(int(m[1]) for p in slide_dir.iterdir() if (m := re.fullmatch(r"slide-(\d+)\.png", p.name)))
    if not pngs:
        raise Invalid(f"no slide-N.png in {slide_dir}")
    entries = load(audit_path)
    if type(entries) is not list:
        raise Invalid("audit.json is not a JSON array")
    by_slide: dict[int, dict] = {}
    for i, e in enumerate(entries):
        bad = [k for k, t in FIELDS.items() if type(e) is not dict or type(e.get(k)) is not t]
        if bad:
            raise Invalid(f"entry {i + 1} {bad[0]} is missing or not {FIELDS[bad[0]].__name__}")
        n = e["slide"]
        if n in by_slide:
            raise Invalid(f"slide {n} has more than one entry")
        if n not in pngs:
            raise Invalid(f"slide {n} has an entry but no slide-{n}.png")
        by_slide[n] = e
    needs_work = advisory_fails = 0
    for n in pngs:
        if n not in by_slide:
            raise Invalid(f"slide {n} has no entry")
        checks, evidence, advisory, verdict = (by_slide[n][k] for k in ("checks", "evidence", "advisory", "verdict"))
        missing = [c for c in RESULTS if c not in checks]
        if missing:
            raise Invalid(f"slide {n} {missing[0]} is missing")
        extra = [c for c in checks if c not in RESULTS]
        if extra:
            raise Invalid(f"slide {n} {extra[0]} is not a check")
        for c, legal in RESULTS.items():
            if checks[c] not in legal:
                raise Invalid(f"slide {n} {c} is {checks[c]!r}, want one of {list(legal)}")
            if checks[c] == "fail" and not (type(evidence.get(c)) is str and evidence[c].strip()):
                raise Invalid(f"slide {n} {c} fails with no evidence")
        if sorted(advisory, key=str) != want_advisory:
            raise Invalid(f"slide {n} advisory {advisory}, want {want_advisory} because fonts.json {why}")
        failing = {c for c in RESULTS if checks[c] == "fail"}
        want_verdict = VERDICT[bool(failing - set(advisory))]
        if verdict != want_verdict:
            raise Invalid(f"slide {n} verdict {verdict}, want {want_verdict}")
        needs_work += verdict == "needs-work"
        advisory_fails += bool(failing & set(advisory))
    return f"AUDIT VALID ({len(pngs)} slides, {needs_work} needs-work, {advisory_fails} advisory)"


def main(argv: list[str]) -> int:
    if len(argv) != 2:
        print("usage: audit_check.py <slide-dir> <audit.json>", file=sys.stderr)
        return 2
    try:
        print(check(Path(argv[0]), Path(argv[1])))
    except Invalid as e:
        print(f"AUDIT INVALID: {e}")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
