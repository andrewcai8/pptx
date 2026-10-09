# The review app

The review app is the second piece of meeting-to-slides. A consultant opens it in a browser, sees what a meeting changes in its deck, decides each change, and gets the final deck. It runs on the consultant's machine and reads and writes files in this repo only.

The app shows slide text, slide pictures, and transcript quotes. It binds to 127.0.0.1, so no other machine can reach it, and it makes no network calls. The confidentiality rules in `CLAUDE.md` still apply to every deck and meeting you open in it, and to anything an AI model reads from its pages or files.

## Start it

Run it from the repo root, because a ChangeSet names its deck by a path relative to the root.

```bash
uv run --project deckcheck review serve
```

It prints `review: http://127.0.0.1:8765`. Open that address in a browser. `--port 0` picks a free port and prints it. Stop the server with Ctrl-C. Only one review server can use a repo at a time, because it is the one writer under `artifacts/review/`. A second one exits and says so.

Processing renders slides, so it needs `soffice`, `pdftoppm`, and `fc-match`, the same tools as `deckcheck render`. A golden scenario also needs its corpus deck. Fetch the corpus decks with `uv run --project deckcheck python .claude/skills/verify-pptx/scripts/corpus.py`.

## Meetings

The home screen lists each folder under `evals/` and `private/meetings/` that holds a `transcript.md` or a `changeset.json`. The title and date come from the transcript's `# ` heading and its `Date: YYYY-MM-DD` line, or else from the ChangeSet.

**Process meeting** runs three steps:

1. The maker writes a ChangeSet.
2. The engine checks it and writes the deck with every change applied.
3. The old deck and the new deck are rendered to slide pictures, about 40 seconds.

If the engine refuses the ChangeSet, the row shows each problem the way `changeset validate` prints it. **Process again** starts over and discards the decisions made so far. If it fails, the row shows why and still opens the earlier review with its decisions.

## Deciding

The left rail shows the new deck. A changed slide carries a count, and a check once every change on it is decided. A deleted slide shows greyed where it used to be.

- A slide whose changes are all text shows the new slide large. Point at a highlight to see the change, its transcript quotes, and the maker's reason. Click to keep the note open.
- A slide with a table, chart, or slide-level change shows the old and new slides side by side and lists every change on it.

Each change takes **Keep new**, **Keep old**, or **Edit myself**. An edit must meet the same rules as the maker's text, and the app shows the engine's message when it does not. Dropping an added slide drops what fills it. The fills you had not decided turn to **Keep old**, and restoring the slide asks about them again. A fill you had decided keeps your decision.

**Apply** runs once every change is decided. It writes the final deck and offers it for download. The deck is byte for byte what `changeset apply` writes from the same ChangeSet. Changing a decision afterwards marks the final deck out of date until you apply again.

**Publish to OneDrive** is not connected. It needs an app that IT has to approve, and it uploads nothing.

## Where files go

Every file the app writes is under `artifacts/review/<evals or private>/<meeting>/`, which git ignores. It never writes to `evals/` or `private/`.

| file | what it holds |
|---|---|
| `changeset.json` | the maker's ChangeSet, then the decisions. The app is its only writer. |
| `executed.pptx` | the deck with every change as the maker wrote it |
| `render/old/`, `render/new/` | slide pictures of the source deck and of `executed.pptx`, with `fonts.json` |
| `final.pptx`, `applied.json` | the last applied deck and the ChangeSet hash it came from |

You can run the engine on the same ChangeSet:

```bash
uv run --project deckcheck changeset apply artifacts/review/evals/solar-market-refresh/changeset.json --out private/solar-final.pptx
```

## The maker

A golden scenario has no maker yet. Its row says **Simulated maker: replays the committed changeset.json**, and processing copies that file. A meeting without a `changeset.json` cannot be processed until the maker exists.

`--maker` runs a command as the maker for every meeting. The app replaces `{meeting}` with the meeting's folder and `{out}` with the path the ChangeSet must be written to. It runs the command without a shell, from the repo root. Its output goes to `maker.log` beside the ChangeSet, and the end of it shows on the home screen when the command fails.

```bash
uv run --project deckcheck review serve --maker 'my-maker {meeting} --out {out}'
```

`deckcheck/src/deckcheck/review/maker.py` holds this seam.

## Prove it in a browser

`deckcheck/scripts/review_proof.mjs` drives the app in headless Chrome. It needs Node 24 and Chrome, and `CHROME` sets the browser path. It starts its own server on a free port and stops it at the end, so stop any other review server first.

```bash
node deckcheck/scripts/review_proof.mjs artifacts/review-proof
```

It processes the insurance, solar, and fmcg golden meetings from the home screen, decides every change through the page, applies, and saves screenshots. For each meeting it runs `changeset apply` on a copy of the committed ChangeSet with the same decisions and checks that both decks have the same bytes. It prints `REVIEW PROOF PASS` or `REVIEW PROOF FAIL` and exits 0 or 1. It starts those three meetings over, so their decisions under `artifacts/review/` are lost.
