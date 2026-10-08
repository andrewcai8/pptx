#!/usr/bin/env bash
# Proves deckcheck catches every house-style rule on known decks and that fix leaves only the report-only ones.
# Evidence survives in artifacts/verify-pptx/<run>/.
set -euo pipefail

ROOT="$(git -C "$(dirname "$0")" rev-parse --show-toplevel)"
RUN_ID="${RUN_ID:-selftest-$(date +%Y%m%d-%H%M%S)-$$}"
EVIDENCE="$ROOT/artifacts/verify-pptx/$RUN_ID"
SCRATCH="$(mktemp -d)"
trap 'rm -rf "$SCRATCH"' EXIT

dc() { (cd "$ROOT" && uv run --quiet --project deckcheck deckcheck "$@"); }
fail() { echo "SELFTEST FAIL: $*" >&2; echo "evidence: $EVIDENCE" >&2; exit 1; }
sha() { python3 -c 'import hashlib,sys; print(hashlib.sha256(open(sys.argv[1], "rb").read()).hexdigest())' "$1"; }

mkdir -p "$EVIDENCE"/{decks,clean,dirty,diff,fix,fixed}

dc doctor > "$EVIDENCE/doctor.txt" || fail "doctor exited $? (see doctor.txt)"

(cd "$ROOT" && uv run --quiet --project deckcheck python deckcheck/scripts/make_sample_decks.py "$SCRATCH") \
	| sed -n '/^expected/,$p' | tail -n +2 | sort > "$EVIDENCE/expected-dirty.txt"
cp "$SCRATCH"/*.pptx "$EVIDENCE/decks/"
dirty_sha="$(sha "$EVIDENCE/decks/dirty.pptx")"

set +e
dc check "$EVIDENCE/decks/clean.pptx" --out "$EVIDENCE/clean" > "$EVIDENCE/clean/stdout.txt" 2>&1; clean_rc=$?
dc check "$EVIDENCE/decks/dirty.pptx" --out "$EVIDENCE/dirty" > "$EVIDENCE/dirty/stdout.txt" 2>&1; dirty_rc=$?
dc diff "$EVIDENCE/decks/clean.pptx" "$EVIDENCE/decks/clean-v2.pptx" --out "$EVIDENCE/diff" > "$EVIDENCE/diff/stdout.txt" 2>&1; diff_rc=$?
dc fix "$EVIDENCE/decks/dirty.pptx" --out "$EVIDENCE/fix/dirty-fixed.pptx" --report "$EVIDENCE/fix" > "$EVIDENCE/fix/stdout.txt" 2>&1; fix_rc=$?
dc check "$EVIDENCE/fix/dirty-fixed.pptx" --out "$EVIDENCE/fixed" > "$EVIDENCE/fixed/stdout.txt" 2>&1; fixed_rc=$?
set -e

[ "$clean_rc" -eq 0 ] || fail "clean.pptx check exited $clean_rc, want 0"
[ "$dirty_rc" -eq 1 ] || fail "dirty.pptx check exited $dirty_rc, want 1"
[ "$diff_rc" -eq 0 ] || fail "diff exited $diff_rc, want 0"
[ "$fix_rc" -eq 1 ] || fail "dirty.pptx fix exited $fix_rc, want 1 (see fix/stdout.txt)"
[ "$fixed_rc" -eq 1 ] || fail "check of the fixed dirty.pptx exited $fixed_rc, want 1"
[ "$(sha "$EVIDENCE/decks/dirty.pptx")" = "$dirty_sha" ] || fail "fix changed its input dirty.pptx"

python3 -c 'import json,sys; [print(v["slide"], v["rule"]) for v in json.load(open(sys.argv[1]))["violations"]]' \
	"$EVIDENCE/dirty/report.json" | sort > "$EVIDENCE/actual-dirty.txt"
diff -u "$EVIDENCE/expected-dirty.txt" "$EVIDENCE/actual-dirty.txt" > "$EVIDENCE/dirty-mismatch.diff" \
	|| fail "dirty.pptx violations differ from expected (see dirty-mismatch.diff)"
rm "$EVIDENCE/dirty-mismatch.diff"

statuses="$(python3 -c 'import json,sys; print(" ".join(f"{s["slide"]}:{s["status"]}" for s in json.load(open(sys.argv[1]))["slides"]))' "$EVIDENCE/diff/diff.json")"
[ "$statuses" = "1:unchanged 2:changed 3:unchanged 4:unchanged 5:added" ] || fail "diff statuses were '$statuses'"

sed -n 's/^fixable rule ids: //p' "$EVIDENCE/doctor.txt" | tr -d ' ' | tr ',' '\n' > "$EVIDENCE/fixable.txt"
awk 'NR == FNR { fixable[$0] = 1; next } !($2 in fixable)' "$EVIDENCE/fixable.txt" "$EVIDENCE/expected-dirty.txt" \
	> "$EVIDENCE/expected-fixed.txt"
python3 -c 'import json,sys; [print(v["slide"], v["rule"]) for v in json.load(open(sys.argv[1]))["violations"]]' \
	"$EVIDENCE/fixed/report.json" | sort > "$EVIDENCE/actual-fixed.txt"
diff -u "$EVIDENCE/expected-fixed.txt" "$EVIDENCE/actual-fixed.txt" > "$EVIDENCE/fixed-mismatch.diff" \
	|| fail "fixed dirty.pptx violations differ from the report-only ones (see fixed-mismatch.diff)"
rm "$EVIDENCE/fixed-mismatch.diff"

echo "SELFTEST PASS ($(wc -l < "$EVIDENCE/expected-dirty.txt" | tr -d ' ') rules caught, diff scoped to slides 2 and 5," \
	"fix left the $(wc -l < "$EVIDENCE/expected-fixed.txt" | tr -d ' ') report-only violations)"
echo "evidence: $EVIDENCE"
