# Fix a deck

`deckcheck fix` reads a deck and writes a new one with the safe house-style fixes applied. It reports every violation it cannot fix safely, so a person can act on it. It never writes to the input file. A later generation step reuses it as the loop that edits a deck and checks it again.

## Sub-features

- `fix-rules` fixes four rules. `no-bullet-end-punctuation` strips the trailing `.`, `;`, or `,`. `min-font-size` raises explicit run sizes below the minimum to the minimum. `within-slide-bounds` moves a text shape or chart back onto the slide, and shrinks it only when it is wider or taller than the slide. `max-fonts-per-slide` replaces the fewest, least-used non-theme fonts with the theme fonts, the major font in titles and the minor font elsewhere.
- `fix-report-only` never fixes `slide-has-title`, `source-on-data-slides`, `title-max-chars`, or `no-placeholder-text`, because a fix would invent content. Each one stays in the report as `report-only`.
- `fix-declined` reports a fixable violation that the fixer judges unsafe, with the reason. Bullets that end in an abbreviation such as `etc.`, `mgmt.`, or `e.g.`, or in an ellipsis such as `....` or `…`, keep their punctuation. Symbol fonts such as Wingdings are never replaced. An oversized table and a shape inside a rotated group are not moved.
- `fix-loop` checks the deck again after each pass, because one fix can expose another. It stops after a pass that fixes nothing, or after three passes that each fixed something. A violation that fires again after its fix is reported as `did-not-stick`, and its change still shows as a `fixed` line. A violation that first appears after the last pass is reported as `pass-limit`.
- `fix-scope` changes only the slides a fixer wrote to. It copies every other entry of the input package byte for byte, in the same order, and replaces only those slides' XML parts. It also checks that reading the deck left every slide unchanged. If not, it prints `error: reading <deck> changed slide N, so fix cannot tell its own edits apart; nothing written` and exits 2.
- `fix-output` prints `PASS` or `FAIL <in> -> <out>: N fixed in P passes, M remain`, without the pass count when nothing was fixed, then one `fixed` line per fix with the value before and after, and one `remains` line per violation left with its reason in brackets, such as `(report-only: fix by hand)` or `(declined: abbreviation etc.)`.
- `fix-report` writes `fix.json` and `fix.md` with `--report`. They hold both decks' sha256, the rules applied, every fix, and every remaining violation.
- `fix-exit` exits 0 when nothing remains, 1 when violations remain, and 2 on a bad deck, bad rules, an `--out` that is the input or a directory, a deck that changed when it was read, or an output it cannot write. The deck is written in both the 0 and 1 cases.
- `doctor` prints `fixable rule ids:` from the same registry `fix` uses.
- `fix-idempotent` copies the input byte for byte when there is nothing to fix, so running `fix` on its own output changes nothing.

## How to get to it (user POV)

- Run `deckcheck fix <deck.pptx> --out <new.pptx>` from anywhere inside the repo. It finds the nearest `standards/house-style.yaml`.
- Add `--rules <file.yaml>` to use another standard, and `--report <dir>` to keep the evidence.

## Driving it with deckcheck

Preconditions:

- Baseline preconditions from `README.md` hold.
- Sample decks exist. Run `uv run --project deckcheck python deckcheck/scripts/make_sample_decks.py $RUN/decks`.

- **Record the input.** Run `shasum -a 256 $RUN/decks/dirty.pptx > $RUN/dirty.sha256`.
- **Fix.** Run `uv run --project deckcheck deckcheck fix $RUN/decks/dirty.pptx --out $RUN/fixed.pptx --report $RUN/fix`. Exit 1. Stdout starts with `FAIL ... 4 fixed in 1 pass, 4 remain`, then lists four `fixed` lines and four `remains ... (report-only: fix by hand)` lines.
- **Only report-only violations remain.** Run `uv run --project deckcheck deckcheck check $RUN/fixed.pptx --out $RUN/fixed-check`. Exit 1. The `(slide, rule)` pairs in `report.json` are `(1, no-placeholder-text)`, `(3, source-on-data-slides)`, `(3, title-max-chars)`, and `(4, slide-has-title)`.
- **The text changes are exactly the fixes.** Run `uv run --project deckcheck deckcheck diff $RUN/decks/dirty.pptx $RUN/fixed.pptx --out $RUN/fix-diff`. Only slide 2 is `changed`, and its diff removes the period from `Three competitors exited the segment in 2025.`. The font, size, and position fixes on slides 2 and 4 do not show in a text diff.
- **The input is untouched.** Run `shasum -a 256 -c $RUN/dirty.sha256` and require `OK`.
- **See the visual fixes.** Run `uv run --project deckcheck deckcheck render $RUN/fixed.pptx --out $RUN/fixed-render` and open `slide-2.png` and `slide-4.png`.
- **Refuse to overwrite the input.** Run `uv run --project deckcheck deckcheck fix $RUN/decks/dirty.pptx --out $RUN/decks/dirty.pptx`. Exit 2, and stderr says `fix never overwrites its input`.
- **Proof.** `selftest.sh` runs the fix, the hash check, and the remaining-pairs check. `corpus.py` runs `fix` on every corpus deck and fails if it changes a slide outside that deck's waived fixable slides.

## Gotchas

- `diff` compares text only. Font, size, and position fixes show as `unchanged`. Prove them with `render` or with the `fixed` lines.
- `within-slide-bounds` measures the box, not the text. A wide label box whose text was already visible still moves flush to the edge, so its text shifts. On the BCG corpus this pushed section labels onto bullets and a heading onto a flag. Check the render of every moved shape.
- A move can land a box on top of another shape, and a larger font can overflow its box. No rule catches either. Check the render.
- The abbreviation list is a fixed set in `fix.py`, seeded from the corpus. An abbreviation that is not on it loses its period. The `fixed` line shows the change.
- Only explicit run sizes and `a:latin` fonts change. Inherited sizes, East Asian and complex-script fonts, and master or layout styles stay as they are.
- `model.py` must read the deck without writing to it, because `fix` writes slide parts from the same tree it read. A python-pptx accessor such as `run.font` or `paragraph.level` adds XML when it is read. If a future read does that, `fix` exits 2 and writes nothing.
