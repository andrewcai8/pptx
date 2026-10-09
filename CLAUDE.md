# pptx

Tools that check, fix, and audit PowerPoint decks against a consulting house style. The long-term goal is in `docs/PLAN.md`: turn a recorded meeting into deck edits a consultant approves with zero changes. What exists today is the quality layer (check, fix, and audit), the edit engine, and the maker, a skill that turns a meeting into a ChangeSet. The review site and the OneDrive publish step are not built yet.

## Confidentiality first

This repo is public. The person using it may work with confidential client decks and data.

- **You are an AI model, so anything you read leaves this machine.** That includes command output, file contents, and images you open. `check`, `diff`, `fix`, and the `changeset` commands print slide titles, bullets, and numbers. `render` makes images you would look at.
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

The edit engine runs a ChangeSet, a JSON file of edits to one deck, against that deck. Its commands print slide text too, so the same rules apply. `docs/changeset.md` describes the format.

```bash
uv run --project deckcheck changeset outline private/meeting/before.pptx --json private/meeting/outline.json
uv run --project deckcheck changeset validate private/meeting/changeset.json
uv run --project deckcheck changeset execute private/meeting/changeset.json --out private/meeting/executed.pptx --review artifacts/meeting/review.json
uv run --project deckcheck changeset apply private/meeting/changeset.json --out private/meeting/final.pptx
```

- `outline` lists every slide, layout, shape, paragraph, table cell, and chart point by the ids a ChangeSet uses.
- `validate` checks every change against the source deck. Exit 1 means it found problems, listed one per change.
- `execute` writes a new deck with every change applied, for review. It never edits the source.
- `apply` writes the final deck from a fresh copy of the source and the review decisions. Exit 1 means some decisions are still pending, and nothing is written.

`uv run --project deckcheck review serve` starts the review app on 127.0.0.1, where a consultant decides each change in a browser. It shows client slides and quotes, so the same rules apply to every page you read from it. `docs/review.md` describes it.

The `verify-pptx` skill in `.claude/skills/verify-pptx/` is the full procedure, including the audit. Follow it when asked to verify, fix, or audit a deck.

The `process-meeting` skill in `.claude/skills/process-meeting/` is the maker. It reads a meeting and writes a ChangeSet, and the engine makes the edits. When asked to turn a meeting into deck changes, ask the ask-first question, then run `meeting process`. It copies the meeting, its data, and its deck into `<out>/stage/` and runs the skill there through `claude -p`, whose file tools reach nothing outside that folder:

```bash
uv run --project deckcheck meeting process private/meetings/q3-steerco --out private/meetings/q3-steerco/run1 --shareable
```

`--shareable` is the ask-first answer. Pass it only when the deck and the meeting may be shared with an AI service under the firm's policy. Exit 0 means the ChangeSet is valid, 1 means it is missing or invalid, 2 means bad input, and 3 means the `claude` CLI is missing or not logged in.

## Adding a real meeting

Real meetings let us measure the meeting-to-slides step on real work. Each meeting is a private scenario in `private/meetings/<name>/`, a folder git ignores. The ask-first rule above applies to every file in it.

Save these for each meeting:

- `before.pptx` is the deck as it stood before the meeting.
- `after.pptx` is the version the consultant actually made after it. It is the best evidence of what the meeting asked for.
- `notes.md` holds the meeting notes or transcript exactly as the recording app exported them.
- `data/` holds any client file a new number came from, such as a spreadsheet.

Then, with the user's permission, help turn the folder into a scenario that `evals/score.py` can read:

- `transcript.md` rewrites the notes as one turn per line, `[00:04:10] Name (Role, Org): text`. Keep `notes.md` as the original.
- `expected.yaml` is the answer key. Draft it from the difference between `before.pptx` and `after.pptx` plus what the notes say, and name the deck as `deck: {file: before.pptx, sha256: <sha256 of before.pptx>}`. Follow the format in `evals/README.md` and copy the shape of a public scenario such as `evals/insurance-workshop-prep/expected.yaml`. Ask the user to confirm each change, each thing discussed that must not change, and each ask too vague to act on.
- Check the key with `GOLDEN_PRIVATE_DIR=private/meetings uv run --project deckcheck python evals/score.py <name> private/meetings/<name>/after.pptx`. The consultant's own `after.pptx` should pass. If it fails, read each reason before changing anything. A change the key missed or got wrong means fixing the key. Two failures mean something else and need a note in `private/feedback.md`, in generic terms. One is an untouched slide reported as changed, which can happen if PowerPoint rewrote it on save. The other is a house-style break the consultant introduced. Do not bend the key to hide either.

Never copy anything from `private/meetings/` into `evals/` or any other tracked folder.

## The rules

`standards/house-style.yaml` holds the rules. They were calibrated against 20 public BCG decks, listed in `.claude/skills/verify-pptx/corpus/known-good.yaml`. Do not edit the rules to make a deck pass.

## Giving feedback

The most useful feedback is a false flag: a slide a consultant would ship as-is that `check` or the audit still flags. For each one, write down:

- the rule or audit check id
- why the slide is fine, in generic terms
- a made-up minimal example that triggers the same flag. `deckcheck/scripts/make_sample_decks.py` shows how to build one with python-pptx.

Keep the notes in `private/feedback.md` and send them to the repo owner by hand. Never put real client content in them.
