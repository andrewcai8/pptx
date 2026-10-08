"""Prove deckcheck passes real BCG decks, except for the waivers in corpus/known-good.yaml.

Usage, from the repo root:
    uv run --project deckcheck python .claude/skills/verify-pptx/scripts/corpus.py [RUN_DIR]

Exit 0 on CORPUS PASS, 1 on a waiver mismatch, 2 on a hash mismatch or unreadable deck, 3 on an unreachable deck.
"""

from __future__ import annotations

import contextlib
import hashlib
import http.client
import json
import os
import sys
import time
import urllib.request
from pathlib import Path

import yaml

from deckcheck.cli import main as deckcheck

ROOT = Path(__file__).resolve().parents[4]
MANIFEST = ROOT / ".claude/skills/verify-pptx/corpus/known-good.yaml"
CACHE = ROOT / "artifacts/verify-pptx/corpus-cache"
RULES = ROOT / "standards/house-style.yaml"
# Some hosts, such as mass.gov, answer 403 unless the request carries a browser's full header set.
BROWSER_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 14_0) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Safari/605.1.15",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
    "Sec-Fetch-Dest": "document",
    "Sec-Fetch-Mode": "navigate",
    "Sec-Fetch-Site": "none",
    "Sec-Fetch-User": "?1",
    "Upgrade-Insecure-Requests": "1",
}


class Unreachable(Exception):
    pass


class BadDeck(Exception):
    pass


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def fetch(deck: dict) -> Path:
    path = CACHE / f"{deck['sha256']}.pptx"
    if path.is_file() and sha256(path) == deck["sha256"]:
        return path
    CACHE.mkdir(parents=True, exist_ok=True)
    partial = path.with_suffix(f".{os.getpid()}.part")
    request = urllib.request.Request(deck["url"], headers=BROWSER_HEADERS)
    try:
        with urllib.request.urlopen(request, timeout=120) as response:
            partial.write_bytes(response.read())
    except (OSError, http.client.HTTPException) as e:
        partial.unlink(missing_ok=True)
        raise Unreachable(str(e)) from e
    got = sha256(partial)
    if got != deck["sha256"]:
        partial.unlink()
        raise BadDeck(f"sha256 {got}, manifest has {deck['sha256']}")
    partial.replace(path)
    print(f"fetched {deck['id']} from {deck['url']}")
    return path


def fired_pairs(deck_path: Path, out: Path) -> set[tuple[int, str]]:
    out.mkdir(parents=True, exist_ok=True)
    with open(out / "stdout.txt", "w") as log, contextlib.redirect_stdout(log), contextlib.redirect_stderr(log):
        code = deckcheck(["check", str(deck_path), "--rules", str(RULES), "--out", str(out)])
    if code not in (0, 1):
        raise BadDeck(f"deckcheck check exited {code} (see {out / 'stdout.txt'})")
    report = json.loads((out / "report.json").read_text())
    return {(v["slide"], v["rule"]) for v in report["violations"]}


def waived_pairs(deck: dict) -> set[tuple[int, str]]:
    return {(slide, w["rule"]) for w in deck.get("waivers", []) for slide in w["slides"]}


def mismatches(deck: dict, fired: set[tuple[int, str]], waived: set[tuple[int, str]]) -> list[str]:
    lines = [f"CORPUS FAIL: {deck['id']} slide {s} [{r}] fires without a waiver" for s, r in sorted(fired - waived)]
    lines += [f"CORPUS FAIL: {deck['id']} slide {s} [{r}] waiver is stale, the rule passes" for s, r in sorted(waived - fired)]
    return lines


def main(argv: list[str]) -> int:
    run = Path(argv[0]) if argv else ROOT / f"artifacts/verify-pptx/corpus-{time.strftime('%Y%m%d-%H%M%S')}-{os.getpid()}"
    decks = yaml.safe_load(MANIFEST.read_text())["decks"]
    failures: list[str] = []
    waived = 0
    for deck in decks:
        try:
            path = fetch(deck)
            fired = fired_pairs(path, run / deck["id"])
        except BadDeck as e:
            print(f"CORPUS FAIL: {deck['id']} {e}")
            print(f"evidence: {run}")
            return 2
        except Unreachable as e:
            print(f"CORPUS INCOMPLETE: {deck['id']} unreachable ({e})")
            print(f"evidence: {run}")
            return 3
        pairs = waived_pairs(deck)
        failures += mismatches(deck, fired, pairs)
        waived += len(pairs)
    for line in failures:
        print(line)
    if not failures:
        print(f"CORPUS PASS ({len(decks)} decks, {waived} waived slide-rule pairs)")
    print(f"evidence: {run}")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
