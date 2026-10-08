# pptx

Tools that check, fix, and audit PowerPoint decks against a consulting house style. The long-term goal is in `docs/PLAN.md`: turn a recorded meeting into deck edits a consultant approves with zero changes. That meeting-to-slides step is not built yet. What exists today is the quality layer: check, fix, and audit.

## Confidentiality first

This repo is public. The person using it may work with confidential client decks and data.

- **You are an AI model, so anything you read leaves this machine.** That includes command output, file contents, and images you open. `check`, `diff`, and `fix` print slide titles, bullets, and numbers. `render` makes images you would look at.
- **Ask before you touch a client deck.** Before you run any command on a deck, read any of its files, or open its renders, ask the user whether that deck may be shared with an AI service under their firm's policy. If the answer is no or unsure, do not run the tools yourself. Give the user the exact commands to run in a separate terminal window, not through Claude Code's `!` prefix, whose output enters the chat. Work only from what they choose to tell you, and skip the audit, which needs you to see the slides. This applies to every step of the `verify-pptx` skill, including opening PNGs and the audit.
- **Keep files out of git.** Keep every deck, data file, transcript, recording, render, and report under `private/` or `artifacts/`. Git ignores both folders. As a backstop, `.gitignore` also blocks common client file types and the tools' report files anywhere in the repo. No list catches every format, so the folder rule is the real guard. Before any commit, check `git status` for anything that came from a client.
- **Never publish client content.** Never commit, push, or open a PR or issue that contains client names, numbers, slide text, screenshots, or file names. If you report a problem upstream, describe it in generic terms and use a made-up example.
- The tools make no network calls. The only exception is `corpus.py`, which downloads public BCG decks.

## Setup (macOS)

```bash
brew install uv
uv sync --project deckcheck
brew install --cask libreoffice      # render and audit only
brew install poppler fontconfig      # render and audit only
uv run --project deckcheck deckcheck doctor
```

Run every command from the repo root, because the tools find `standards/house-style.yaml` from there. `doctor` must exit 0. `render` and the audit need `soffice`, `pdftoppm`, and `fc-match` all present. Install the deck's real fonts so renders wrap text the way PowerPoint does. If `fonts.json` shows a font as substituted, install it or a metric-compatible stand-in, such as Carlito for Calibri.

## Use

Copy the deck into `private/` first. Never run the tools on a file inside a synced OneDrive or SharePoint folder.

```bash
uv run --project deckcheck deckcheck check private/deck.pptx --out artifacts/run1/check
uv run --project deckcheck deckcheck fix private/deck.pptx --out private/deck-fixed.pptx --report artifacts/run1/fix
uv run --project deckcheck deckcheck diff private/deck.pptx private/deck-fixed.pptx --out artifacts/run1/diff
uv run --project deckcheck deckcheck render private/deck-fixed.pptx --out artifacts/run1/render
```

- `check` lists every house-style violation with its slide and rule. Exit 1 means it found violations, which is normal.
- `fix` writes a new deck. It never edits the input. It fixes bullet end punctuation, text below the minimum size, and extra fonts, and lists everything else for a person. Exit 1 means some violations remain for a person.
- `diff` compares slide text. It does not see font, size, color, or position changes, so it can call a slide `fix` changed `unchanged`. Compare renders to see those.
- `render` writes slide PNGs, a PDF copy of the deck, and `fonts.json`, which lists any font it had to substitute.

The `verify-pptx` skill in `.claude/skills/verify-pptx/` is the full procedure, including the audit. Follow it when asked to verify, fix, or audit a deck.

## The rules

`standards/house-style.yaml` holds the rules. They were calibrated against 20 public BCG decks, listed in `.claude/skills/verify-pptx/corpus/known-good.yaml`. Do not edit the rules to make a deck pass.

## Giving feedback

The most useful feedback is a false flag: a slide a consultant would ship as-is that `check` or the audit still flags. For each one, write down:

- the rule or audit check id
- why the slide is fine, in generic terms
- a made-up minimal example that triggers the same flag. `deckcheck/scripts/make_sample_decks.py` shows how to build one with python-pptx.

Keep the notes in `private/feedback.md` and send them to the repo owner by hand. Never put real client content in them.
