# Plan for meeting-to-slides automation

Consultants lose hours formatting PowerPoint between meetings. This tool takes a recorded meeting and produces the deck updates that came out of it. The consultant reviews and approves those changes, and the tool then publishes them to the client's OneDrive. The goal is a deck the consultant approves with zero edits. This document calls that a one-shot deck.

## What the consultant sees

1. The consultant records a meeting. The recorder posts the transcript to the tool when the meeting ends.
2. The tool lists the changes it took from the meeting, for example "update the market sizing on slide 4 of `Q3 Steering.pptx`".
3. The tool edits copies of those decks in a staging area. It never writes to the client OneDrive at this step.
4. The consultant opens the review site. Each changed slide shows up to four options next to the original.
5. The consultant picks an option per slide, or edits one, and approves the deck.
6. The tool uploads the approved deck to OneDrive as a new version of the original file. OneDrive keeps version history, so a bad publish can be rolled back.

## One-shot is a measurable bar

The approve-with-zero-edits rate is the number this project climbs. A deck must clear four automated gates before a human sees it:

- **House style.** `deckcheck check` passes against `standards/house-style.yaml`. The file encodes company rules, such as at most three fonts on a slide and no end punctuation on bullets.
- **Scope.** `deckcheck diff` against the source deck shows changes only on the slides the meeting asked for.
- **Provenance.** Every number on a changed slide traces to a client file and cell, or to a URL with a retrieval date. Market data that cannot be cited does not go on a slide.
- **Intent.** The requested change is present. Golden scenarios check this (phase 2).

The first two gates exist today in `deckcheck/`. The `verify-pptx` skill in `.claude/skills/` shows an agent how to run them and capture evidence.

## Data flow

1. **Capture.** A recorder produces the transcript. Granola and Fireflies both have connectors available to this workspace.
2. **Plan the changes.** Claude reads the transcript and an index of the client folder. It returns a typed change list with one entry per change, holding the deck path, the slide, the intent, and the data needed.
3. **Gather data.** Client files come from OneDrive through Microsoft Graph with read-only scope. Market data comes from Parallel or Exa through their APIs. Each fact is stored with its URL, the quoted passage, and the retrieval date in an Excel backup workbook next to the deck.
4. **Edit.** The tool copies the source deck to staging and edits the copy with `python-pptx`, reusing the deck's own layouts and master. Editing the client's file keeps the client format. Building from a blank template would lose it.
5. **Verify.** `deckcheck check` and `deckcheck diff` run on the copy. Violations go back to step 4 as feedback, up to a fixed retry limit. `deckcheck render` then produces slide images for the review site.
6. **Review and publish.** These steps happen on the review site, as described above. Publishing needs the Graph write scope, and only the approve action uses it.

## think-cell needs a Windows worker

think-cell automation fills think-cell charts in a template from Excel data (the `PresentationFromTemplate` API) or from JSON (`.ppttc` files). Both paths run inside a machine with PowerPoint and a licensed think-cell install. See [Advanced report automation](https://www.think-cell.com/en/resources/manual/introductionautomation) and [Automation with JSON data](https://think-cell.com/en/resources/manual/jsondataautomation). A cloud-only pipeline cannot produce think-cell charts, so this plan adds a Windows worker in phase 6.

Until then, charts are native PowerPoint charts, and their data lives in the Excel backup workbook from day one. When the worker lands, adding think-cell means rendering that same workbook through think-cell. The data stays where it is.

## Phases

Each phase ends in a check that `deckcheck` or the golden set can run.

| Phase | Builds | Done when |
| --- | --- | --- |
| 0 | `deckcheck`, `standards/house-style.yaml`, `verify-pptx` skill | Clean sample deck passes and dirty sample deck fails with the exact expected violations. |
| 1 | Fix mode. Takes a deck and writes a new deck with house-style violations fixed. | Output passes `check`. `diff` shows no text change beyond the fixes. |
| 2 | Golden scenarios in `evals/<name>/` holding a transcript, input decks, and `expected.yaml` | A runner scores the one-shot rate across the set. |
| 3 | Transcript to change list to edited decks in local staging | The golden-set score rises above the phase 1 baseline. |
| 4 | Market-data slides from Parallel or Exa, with citations | Every number on a generated slide resolves to a stored source. |
| 5 | Review site, OneDrive staging, and publish through Graph | A consultant approves a deck end to end on a test tenant. |
| 6 | Windows worker for think-cell | A think-cell chart in a client template updates from the backup workbook. |

Phase 1 comes first because it needs no connectors. It also tests the edit-and-verify loop that every later phase reuses.

## Repo layout

- `standards/house-style.yaml` holds the company slide rules. The generator and `deckcheck` both read it.
- `deckcheck/` is the Python CLI that checks, diffs, and renders decks.
- `.claude/skills/verify-pptx/` holds the verification skill and its feature map.
- `docs/PLAN.md` is this document.

## Open decisions

These need an answer from the team. No experiment can settle them.

- Which meeting recorder the consultants use. Granola, Fireflies, and the Teams built-in recorder each need a different trigger.
- Whether IT will grant admin consent for a Microsoft Graph app on the client tenant, and with which scopes.
- Where staging lives. A separate OneDrive folder needs no new storage. App storage keeps drafts out of the client tenant.
- The company template and 10 to 20 exemplar decks. The house style currently encodes the two rules in the brief plus common consulting defaults. Real decks would show which rules the firm actually follows. Until they arrive, the guessed defaults are calibrated against five public BCG decks in the `verify-pptx` known-good corpus.
- What "no bullet points on the end of a sentence" means. `deckcheck` currently reads it as "a bullet must not end in `.`, `;`, or `,`".
- Whether a Windows machine with Office and think-cell is available for phase 6.
