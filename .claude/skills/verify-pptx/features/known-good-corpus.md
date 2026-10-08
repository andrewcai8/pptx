# Pass the known-good corpus

`corpus.py` runs `deckcheck check` on five real BCG decks and proves the house style accepts them. Every violation that still fires must match a waiver in `corpus/known-good.yaml`, and every waiver must still fire. A rule change that flags good decks, or one that quietly stops catching a waived case, fails the gate.

## Sub-features

- `corpus-fetch` downloads each deck into `artifacts/verify-pptx/corpus-cache/<sha256>.pptx` and reuses it on later runs. The decks remain BCG's property, so the repo stores only their URLs and hashes.
- `corpus-pass` prints `CORPUS PASS (5 decks, M waived slide-rule pairs)` and exits 0.
- `corpus-mismatch` prints one `CORPUS FAIL: <id> slide N [rule] ...` line per unwaived or stale pair and exits 1.
- `corpus-hash` prints `CORPUS FAIL: <id> sha256 ...` and exits 2 when a download does not match the manifest.
- `corpus-unreachable` prints `CORPUS INCOMPLETE: <id> unreachable (<error>)` and exits 3.
- `corpus-evidence` writes `report.json`, `outline.md`, and `stdout.txt` per deck under the run directory.

## How to get to it (user POV)

- Run `uv run --project deckcheck python .claude/skills/verify-pptx/scripts/corpus.py` from the repo root after changing `standards/house-style.yaml` or `deckcheck/src/deckcheck/rules.py`.
- Pass a directory as the only argument to choose the evidence path.

## Driving it with deckcheck

Preconditions:

- Baseline preconditions from `README.md` hold.
- The machine can reach the five URLs in `corpus/known-good.yaml`, or the cache already holds the decks.

- **Pass.** Run the script. Exit 0 and the last two lines are `CORPUS PASS (...)` and `evidence: <dir>`. A second run prints no `fetched` lines because it reuses the cache.
- **Unwaived violation.** Copy the manifest aside, remove one slide from a waiver, and rerun. Exit 1 with `CORPUS FAIL: <id> slide N [rule] fires without a waiver`. Restore the manifest.
- **Stale waiver.** Add a slide that does not fire to a waiver and rerun. Exit 1 with `waiver is stale, the rule passes`. Restore the manifest.
- **Read a deck's result.** Open `<dir>/<id>/report.json` for the violations and `<dir>/<id>/outline.md` for the text deckcheck saw.

## Gotchas

- A waiver records a deliberate difference between BCG practice and what the pipeline needs. Loosening a rule to clear a waiver changes what the pipeline accepts, so it needs a reason from the brief, not from the corpus.
- The four other `.pptx` links on the slideworks page are 404 or bot-walled, so the corpus holds five decks.
- A host can change or remove a deck. A new hash fails the gate with exit 2, and a dead link reports `CORPUS INCOMPLETE`. Never count either as a pass.
- Fix a new violation in deckcheck when the render shows the slide is fine. Waive it only when BCG practice deliberately differs from the house style.
