# Deck verification map

This directory is the maintained source for verifying deck output in this repo. Read this index, then use the matching feature file as the recipe.

## Baseline preconditions

- Run every command from the repo root.
- `uv sync --project deckcheck` has run.
- `uv run --project deckcheck deckcheck doctor` exits 0 and names `standards/house-style.yaml`.
- `RUN=artifacts/verify-pptx/<unique-id>` is set, and every `--out` path sits under it.
- The input decks are copies in `$RUN/`, never files in a synced OneDrive folder.

## Driving conventions

- Prefix every command with `uv run --project deckcheck`.
- Assert on exit codes and on the JSON reports, not on wording in stdout.
- Treat `standards/house-style.yaml` as read-only during a proof.

## Proof and skip reporting

- Record the command, the exit code, and the report path for each step.
- Record the deck sha256 from `report.json` or `diff.json`, so the proof names one exact file.
- Report `render` and the audit as skipped when `doctor` prints `soffice`, `pdftoppm`, or `fc-match` as `missing`. Never report it as passed through `check`.
- Do not report a deck as verified when only the sample decks were checked.

## Feature entry contract

Each feature file starts with an H1 title and one paragraph describing the behavior. It then has exactly four H2 sections in this order: `Sub-features`, `How to get to it (user POV)`, `Driving it with deckcheck`, and `Gotchas`.

## Features

- [Check the house style](./check-house-style.md) covers the eight rules, the reports, and config errors.
- [Diff against the source deck](./diff-decks.md) covers scoping a change to the requested slides and proving the source is untouched.
- [Fix a deck](./fix-deck.md) covers writing a copy with the three fixable rules repaired and the rest reported.
- [Render slides](./render-slides.md) covers PNG previews through LibreOffice and the missing-tool path.
- [Audit slides](./audit-slides.md) covers the advisory review of rendered slides against the rubric, the font report, and `audit_check.py`.
- [Pass the known-good corpus](./known-good-corpus.md) covers the real BCG decks that the house style must accept and their waivers.
