# Check the house style

`deckcheck check` reads a deck and the company standard in `standards/house-style.yaml`, and reports every slide that breaks a rule. A consultant should never see a deck that fails this check.

## Sub-features

- `check-pass` prints `PASS <deck> (N slides, M rules)` and exits 0 for a compliant deck.
- `check-fail` prints `FAIL <deck>: K violations`, one line per violation, and exits 1.
- `check-report` writes `report.json` and `outline.md` with `--out`.
- `check-rules` enforces `max-fonts-per-slide`, `no-bullet-end-punctuation`, `slide-has-title`, `title-max-chars`, `min-font-size`, `no-placeholder-text`, `within-slide-bounds`, and `source-on-data-slides`.
- `check-config-error` exits 2 on an unknown rule id, a missing or unknown param, or an invalid regex.
- `check-bad-deck` exits 2 on a file that is not a .pptx.

## How to get to it (user POV)

- Run `deckcheck check <deck.pptx>` from anywhere inside the repo. It finds the nearest `standards/house-style.yaml`.
- Run `deckcheck check <deck.pptx> --rules <file.yaml>` to use another standard.

## Driving it with deckcheck

Preconditions:

- Baseline preconditions from `README.md` hold.
- Sample decks exist. Run `uv run --project deckcheck python deckcheck/scripts/make_sample_decks.py $RUN/decks`.

- **Pass.** Run `uv run --project deckcheck deckcheck check $RUN/decks/clean.pptx --out $RUN/clean`. Exit 0, stdout starts with `PASS`, and `$RUN/clean/report.json` has `"passed": true`.
- **Fail with every rule.** Run `uv run --project deckcheck deckcheck check $RUN/decks/dirty.pptx --out $RUN/dirty`. Exit 1. The `(slide, rule)` pairs in `report.json` equal the list that `make_sample_decks.py` printed.
- **Read the content.** Open `$RUN/clean/outline.md`. Each slide shows its title, each paragraph, and its resolved fonts.
- **Config error.** Write a copy of the rules with an extra key `not-a-rule: {}` to `$RUN/bad-rules.yaml`. Run `uv run --project deckcheck deckcheck check $RUN/decks/clean.pptx --rules $RUN/bad-rules.yaml`. Exit 2 and stderr names `'not-a-rule'`.
- **Bad deck.** Run `uv run --project deckcheck deckcheck check deckcheck/pyproject.toml`. Exit 2 and stderr says `cannot read deck`.
- **Proof.** `.claude/skills/verify-pptx/scripts/selftest.sh` runs the pass and fail steps and asserts the exact pairs.

## Gotchas

- Font resolution reads explicit run fonts and the theme major and minor fonts. It ignores fonts set in master or layout text styles, so real templates can count one font twice under two names.
- Body and object placeholders count as bullets unless the paragraph sets `buNone`. Templates that set `buNone` in the master get false bullet hits.
- `min-font-size` sees only explicit run sizes. Inherited sizes and shrink-to-fit text are not checked.
- `source-on-data-slides` matches the prefix case-sensitively and detects only native charts and tables. A pasted chart image or a think-cell chart passes unseen.
- `within-slide-bounds` ignores rotation and flags intentional full-bleed shapes.
