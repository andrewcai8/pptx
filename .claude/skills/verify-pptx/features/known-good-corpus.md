# Pass the known-good corpus

`corpus.py` runs `deckcheck check` on twenty real BCG decks and proves the house style accepts them. Every violation that still fires must match a waiver in `corpus/known-good.yaml`, and every waiver must still fire. A rule change that flags good decks, or one that quietly stops catching a waived case, fails the gate.

## Sub-features

- `corpus-fetch` downloads each deck into `artifacts/verify-pptx/corpus-cache/<sha256>.pptx` and reuses it on later runs. It sends a browser's full header set, because some hosts refuse a bare user agent. The decks remain BCG's property, so the repo stores only their URLs and hashes.
- `corpus-pass` prints `CORPUS PASS (20 decks, M waived slide-rule pairs, K fixed in scope)` and exits 0.
- `corpus-fix-scope` runs `deckcheck fix` on each deck and compares the two packages entry by entry. It prints a `CORPUS FAIL: <id> fix changed ...` line and exits 1 when an entry was added or removed, or when a changed entry is not the XML part of one of the deck's waived slides for a fixable rule. It also fails when `fix` exits 2 or the cached input's hash changes.
- `corpus-mismatch` prints one `CORPUS FAIL: <id> slide N [rule] ...` line per unwaived or stale pair and exits 1.
- `corpus-hash` prints `CORPUS FAIL: <id> sha256 ...` and exits 2 when a download does not match the manifest.
- `corpus-unreachable` prints `CORPUS INCOMPLETE: <id> unreachable (<error>)` and exits 3.
- `corpus-evidence` writes `report.json`, `outline.md`, and `stdout.txt` per deck under the run directory, plus `fixed.pptx`, `fix.txt`, and `fix/` from the fix run.

## How to get to it (user POV)

- Run `uv run --project deckcheck python .claude/skills/verify-pptx/scripts/corpus.py` from the repo root after changing `standards/house-style.yaml` or any file in `deckcheck/src/deckcheck/`.
- Pass a directory as the only argument to choose the evidence path.

## Driving it with deckcheck

Preconditions:

- Baseline preconditions from `README.md` hold.
- The machine can reach the twenty URLs in `corpus/known-good.yaml`, or the cache already holds the decks.

- **Pass.** Run the script. Exit 0 and the last two lines are `CORPUS PASS (...)` and `evidence: <dir>`. A second run prints no `fetched` lines because it reuses the cache.
- **Unwaived violation.** Copy the manifest aside, remove one slide from a waiver, and rerun. Exit 1 with `CORPUS FAIL: <id> slide N [rule] fires without a waiver`. Restore the manifest.
- **Stale waiver.** Add a slide that does not fire to a waiver and rerun. Exit 1 with `waiver is stale, the rule passes`. Restore the manifest.
- **Read a deck's result.** Open `<dir>/<id>/report.json` for the violations and `<dir>/<id>/outline.md` for the text deckcheck saw.

## Gotchas

- A waiver records a deliberate difference between BCG practice and what the pipeline needs. Loosening a rule to clear a waiver changes what the pipeline accepts, so it needs a reason from the brief, not from the corpus.
- Five decks come from the slideworks list. Its four other `.pptx` links are 404 or bot-walled. The other fifteen were found by search, and the manifest groups the decks by source.
- Five decks hid whole classes of false positives, such as rotated axis labels and copied layout names. Grow the corpus before you trust a rule change that only five decks support.
- A deck whose host went dead is pinned to a Wayback snapshot, a `web.archive.org/web/<timestamp>id_/<url>` URL. The `id_` suffix returns the original bytes, so the hash still matches.
- mass.gov answers 403 to a request without `Accept`, `Accept-Language`, and `Sec-Fetch-*` headers. Keep the full browser header set in `corpus.py`.
- The fix run takes about 25 seconds on top of the check, so a full corpus run takes about 35 seconds.
- The two RCC Road to Recovery compilations are left out because they duplicate the chapter decks.
- A host can change or remove a deck. A new hash fails the gate with exit 2, and a dead link reports `CORPUS INCOMPLETE`. Never count either as a pass.
- Fix a new violation in deckcheck when the render shows the slide is fine. Waive it only when BCG practice deliberately differs from the house style.
