# pptx

Tools that check, fix, and audit PowerPoint decks against a consulting house style. The long-term goal is in `docs/PLAN.md`: turn a recorded meeting into deck edits a consultant approves with zero changes. That meeting-to-slides step is not built yet. What exists today is the quality layer: check, fix, and audit.

## Confidentiality first

This repo is public. The person using it may work with confidential client decks and data.

- Keep every deck, data file, transcript, and output under `private/` or `artifacts/`. Both are gitignored, and so is every `.pptx` file.
- Never commit, push, or open a PR or issue that contains client names, numbers, slide text, screenshots, or file names. If you report a problem upstream, describe it in generic terms and use a made-up example.
- `check`, `fix`, `diff`, and `render` run locally. No deck content leaves the machine.
- The audit step sends slide images to an AI model. Before auditing a deck, ask the user whether that deck may be shared with an AI service under their firm's policy. If they are unsure, skip the audit.

## Setup (macOS)

```bash
brew install uv
uv sync --project deckcheck
brew install --cask libreoffice      # render and audit only
brew install poppler fontconfig      # render and audit only
uv run --project deckcheck deckcheck doctor
```

`doctor` must exit 0. `render` and the audit need `soffice`, `pdftoppm`, and `fc-match` all present. Install the deck's real fonts (macOS ships Trebuchet MS, Microsoft Office ships Calibri) so renders wrap text the way PowerPoint does.

## Use

Copy the deck into `private/` first. Never run the tools on a file inside a synced OneDrive or SharePoint folder.

```bash
uv run --project deckcheck deckcheck check private/deck.pptx --out artifacts/run1/check
uv run --project deckcheck deckcheck fix private/deck.pptx --out private/deck-fixed.pptx --report artifacts/run1/fix
uv run --project deckcheck deckcheck diff private/deck.pptx private/deck-fixed.pptx --out artifacts/run1/diff
uv run --project deckcheck deckcheck render private/deck-fixed.pptx --out artifacts/run1/render
```

- `check` lists every house-style violation with its slide and rule.
- `fix` writes a new deck. It never edits the input. It fixes bullet end punctuation, text below the minimum size, and extra fonts, and lists everything else for a person.
- `diff` shows which slides changed and how.
- `render` writes slide PNGs and `fonts.json`, which lists any font it had to substitute.

The `verify-pptx` skill in `.claude/skills/verify-pptx/` is the full procedure, including the audit. Follow it when asked to verify, fix, or audit a deck.

## The rules

`standards/house-style.yaml` holds the rules. They were calibrated against 20 public BCG decks, listed in `.claude/skills/verify-pptx/corpus/known-good.yaml`. Do not edit the rules to make a deck pass.

## Giving feedback

The most useful feedback is a false flag: a slide a consultant would ship as-is that `check` or the audit still flags. For each one, write down:

- the rule or audit check id
- why the slide is fine, in generic terms
- a made-up minimal example that triggers the same flag

Keep the notes in `private/feedback.md` and share them by hand. Never put real client content in them.
