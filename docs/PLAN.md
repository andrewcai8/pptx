# Meeting-to-slides automation

Consultants lose hours formatting PowerPoint between meetings. This tool takes a recorded meeting and produces the deck updates that came out of it. The consultant reviews and approves those changes, and the tool then publishes them to the client's OneDrive. The goal is a deck the consultant approves with zero edits. This document calls that a one-shot deck. It records the goal and how we measure it, not a fixed order of work.

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
- **Intent.** The requested change is present. Example meetings with known expected edits check this.

The first two gates exist today in `deckcheck/`. The `verify-pptx` skill in `.claude/skills/` shows an agent how to run them and capture evidence. The golden scenarios in `evals/` check scope and the required facts by script. They list the intent checks, and no intent checker runs them yet. Provenance has no automated gate yet.

## Status

Steps 2 to 5 of "What the consultant sees" run today on one Mac, from meeting notes the consultant saves by hand. The review app offers one proposed edit per change, not four, and the consultant keeps it, reverts it, or edits it. `CLAUDE.md` has the procedure under "Run a meeting end to end".

- `meeting process` runs the `process-meeting` skill through `claude -p` and writes a ChangeSet. It has run only against a fake `claude`.
- The edit engine checks a ChangeSet against its deck and writes the edited deck. It never edits the source.
- The review app shows each change, records the consultant's decisions, and writes the final deck.

Two steps are not built yet. Step 1 needs a hook from the meeting recorder. Step 6, the OneDrive upload, needs IT approval. `CLAUDE.md` lists the known gaps.

## Repo layout

- `standards/house-style.yaml` holds the company slide rules. `deckcheck` reads it.
- `deckcheck/` is the Python package. It has these commands:
  - `deckcheck` checks, fixes, diffs, and renders decks.
  - `changeset` outlines a deck, and validates, executes, and applies a ChangeSet. `docs/changeset.md` describes the format.
  - `meeting process` runs the maker on one meeting folder.
  - `review serve` runs the review app. `docs/review.md` describes it.
- `.claude/skills/verify-pptx/` holds the verification skill and its feature map.
- `.claude/skills/process-meeting/` holds the maker skill.
- `evals/` holds the golden meeting scenarios and the script that scores a deck against them.
- `private/` and `artifacts/` hold client files and run outputs. Git ignores both.
- `docs/PLAN.md` is this document.

## Open decisions

These need an answer from the team. No experiment can settle them.

- Which meeting recorder the consultants use. Granola, Fireflies, and the Teams built-in recorder each need a different trigger.
- Whether IT will grant admin consent for a Microsoft Graph app on the client tenant, and with which scopes.
- Where staging lives. A separate OneDrive folder needs no new storage. App storage keeps drafts out of the client tenant.
- The company template and 10 to 20 exemplar decks. The house style currently encodes the two rules in the brief plus common consulting defaults. Real decks would show which rules the firm actually follows. Until they arrive, the guessed defaults are calibrated against twenty public BCG decks in the `verify-pptx` known-good corpus.
- What "no bullet points on the end of a sentence" means. `deckcheck` currently reads it as "a bullet must not end in `.`, `;`, or `,`".
- Whether a Windows machine with Office and think-cell is available. think-cell automation only runs inside PowerPoint with a licensed install, so think-cell charts need one.
