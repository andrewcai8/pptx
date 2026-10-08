# Diff against the source deck

`deckcheck diff` compares a source deck with the pipeline's new deck slide by slide. It proves the change touched only the slides the meeting asked for.

## Sub-features

- `diff-status` labels each slide index `unchanged`, `changed`, `added`, or `removed`.
- `diff-text` prints a unified diff of each changed slide's text, one line per paragraph prefixed by its shape name.
- `diff-hashes` prints the sha256 of both decks.
- `diff-report` writes `diff.json` and `diff.md` with `--out`.

## How to get to it (user POV)

- Run `deckcheck diff <source.pptx> <new.pptx>`.

## Driving it with deckcheck

Preconditions:

- Baseline preconditions from `README.md` hold.
- Sample decks exist in `$RUN/decks` from `make_sample_decks.py`.

- **Scope a change.** Run `uv run --project deckcheck deckcheck diff $RUN/decks/clean.pptx $RUN/decks/clean-v2.pptx --out $RUN/diff`. Exit 0. Stdout lists `slide 2: changed` and `slide 5: added` and no other slide.
- **Read the change.** Open `$RUN/diff/diff.md`. Slide 2 shows `-Content Placeholder 2: Enterprise renewals flat at 94%` and `+Content Placeholder 2: Enterprise renewals flat at 95%`.
- **No change.** Run `uv run --project deckcheck deckcheck diff $RUN/decks/clean.pptx $RUN/decks/clean.pptx`. Stdout says `no slide changes`.
- **Source untouched.** Before the pipeline runs, run `shasum -a 256 <source.pptx> > $RUN/source.sha256`. Afterwards, run `shasum -a 256 -c $RUN/source.sha256` and require `OK`.

## Gotchas

- Slides are matched by index. Inserting a slide in the middle marks every later slide `changed`. Read the text diff before calling a change out of scope.
- Only text is compared. Font, color, position, and chart data changes show as `unchanged`.
- Table cell text is included. Chart data is not.
