#!/usr/bin/env bash
# Proves deckcheck catches every house-style rule on known decks. Evidence survives in artifacts/verify-pptx/<run>/.
set -euo pipefail

ROOT="$(git -C "$(dirname "$0")" rev-parse --show-toplevel)"
RUN_ID="${RUN_ID:-selftest-$(date +%Y%m%d-%H%M%S)-$$}"
EVIDENCE="$ROOT/artifacts/verify-pptx/$RUN_ID"
SCRATCH="$(mktemp -d)"
trap 'rm -rf "$SCRATCH"' EXIT

dc() { (cd "$ROOT" && uv run --quiet --project deckcheck deckcheck "$@"); }
fail() { echo "SELFTEST FAIL: $*" >&2; echo "evidence: $EVIDENCE" >&2; exit 1; }

mkdir -p "$EVIDENCE"/{decks,clean,dirty,diff}

dc doctor > "$EVIDENCE/doctor.txt" || fail "doctor exited $? (see doctor.txt)"

(cd "$ROOT" && uv run --quiet --project deckcheck python deckcheck/scripts/make_sample_decks.py "$SCRATCH") \
	| sed -n '/^expected/,$p' | tail -n +2 | sort > "$EVIDENCE/expected-dirty.txt"
cp "$SCRATCH"/*.pptx "$EVIDENCE/decks/"

set +e
dc check "$EVIDENCE/decks/clean.pptx" --out "$EVIDENCE/clean" > "$EVIDENCE/clean/stdout.txt" 2>&1; clean_rc=$?
dc check "$EVIDENCE/decks/dirty.pptx" --out "$EVIDENCE/dirty" > "$EVIDENCE/dirty/stdout.txt" 2>&1; dirty_rc=$?
dc diff "$EVIDENCE/decks/clean.pptx" "$EVIDENCE/decks/clean-v2.pptx" --out "$EVIDENCE/diff" > "$EVIDENCE/diff/stdout.txt" 2>&1; diff_rc=$?
set -e

[ "$clean_rc" -eq 0 ] || fail "clean.pptx check exited $clean_rc, want 0"
[ "$dirty_rc" -eq 1 ] || fail "dirty.pptx check exited $dirty_rc, want 1"
[ "$diff_rc" -eq 0 ] || fail "diff exited $diff_rc, want 0"

python3 -c 'import json,sys; [print(v["slide"], v["rule"]) for v in json.load(open(sys.argv[1]))["violations"]]' \
	"$EVIDENCE/dirty/report.json" | sort > "$EVIDENCE/actual-dirty.txt"
diff -u "$EVIDENCE/expected-dirty.txt" "$EVIDENCE/actual-dirty.txt" > "$EVIDENCE/dirty-mismatch.diff" \
	|| fail "dirty.pptx violations differ from expected (see dirty-mismatch.diff)"
rm "$EVIDENCE/dirty-mismatch.diff"

statuses="$(python3 -c 'import json,sys; print(" ".join(f"{s["slide"]}:{s["status"]}" for s in json.load(open(sys.argv[1]))["slides"]))' "$EVIDENCE/diff/diff.json")"
[ "$statuses" = "1:unchanged 2:changed 3:unchanged 4:unchanged 5:added" ] || fail "diff statuses were '$statuses'"

echo "SELFTEST PASS ($(wc -l < "$EVIDENCE/expected-dirty.txt" | tr -d ' ') rules caught, diff scoped to slides 2 and 5)"
echo "evidence: $EVIDENCE"
