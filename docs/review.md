# The review app

The review app is the second piece of meeting-to-slides. A consultant opens it in a browser, sees what a meeting changes in its deck, decides each change, and gets the final deck. It runs on the consultant's machine and reads and writes files in this repo only.

The app shows slide text, slide pictures, and transcript quotes. It binds to 127.0.0.1, so no other machine can reach it, and it makes no network calls. The confidentiality rules in `CLAUDE.md` still apply to every deck and meeting you open in it, and to anything an AI model reads from its pages or files.

## Start it

Run it from the repo root, because a ChangeSet names its deck by a path relative to the root.

```bash
uv run --project deckcheck review serve
```

It prints `review: http://127.0.0.1:8765`. Open that address in a browser. `--port 0` picks a free port and prints it. Stop the server with Ctrl-C. Only one review server can use a repo at a time, because it is the one writer under `artifacts/review/`. A second one exits and says so.

Processing renders slides, so it needs `soffice`, `pdftoppm`, and `fc-match`, the same tools as `deckcheck render`. A golden scenario also needs its corpus deck. Fetch the corpus decks with `uv run --project deckcheck python .claude/skills/verify-pptx/scripts/corpus.py`. A private meeting needs its deck at `private/meetings/<name>/before.pptx`.

## Meetings

The home screen lists each folder under `evals/` and `private/meetings/` that holds a `transcript.md`, a `notes.md`, or a `changeset.json`. Any folder name works, spaces included. It skips names that start with `.` or `_`, and names that end in `.partial` or `.discard`, which the app uses for its own work. The title and date come from the first `# ` heading and `Date: YYYY-MM-DD` line in `transcript.md`, else `notes.md`, else the ChangeSet. Without a heading, the title is the folder name.

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
| `maker.log` | the output of the last `--maker` run, kept when it fails |

You can run the engine on the same ChangeSet:

```bash
uv run --project deckcheck changeset apply artifacts/review/evals/solar-market-refresh/changeset.json --out private/solar-final.pptx
```

## The maker

The maker writes the ChangeSet a review starts from. A golden scenario keeps its committed `changeset.json`, so its row says **Simulated maker: replays the committed changeset.json**, and processing copies that file. A private meeting that holds its own `changeset.json` replays it the same way, and its row says **Replays changeset.json**.

A meeting without a `changeset.json` needs `--maker`, a command the app runs for each such meeting. The app replaces `{meeting}` with the meeting's folder, `{out}` with the path the ChangeSet must be written to, and `{dir}` with the folder that holds `{out}`. It runs the command without a shell, from the repo root. Its output goes to `maker.log` beside the ChangeSet. When the command fails, the end of it shows on the home screen, and the row links to the whole `maker.log`, which the app keeps. Without `--maker`, such a meeting cannot be processed.

To run the real maker, `meeting process`:

```bash
uv run --project deckcheck review serve --maker 'uv run --project deckcheck meeting process {meeting} --out {dir} --shareable'
```

`--shareable` tells `meeting process` that the user answered the ask-first question in `CLAUDE.md`. The app still asks for each meeting. When a meeting without a `changeset.json` goes to the `--maker` command, **Process meeting** and **Process again** first show this confirm:

> This sends the meeting notes, data and deck to Claude (Anthropic). Only continue if your firm allows sharing this deck with an AI service.

Cancel sends nothing. The server refuses to run the maker unless the page sends that consent. A golden scenario, or any meeting that replays its own `changeset.json`, sends nothing to an AI service, so it skips the confirm.

`deckcheck/src/deckcheck/review/maker.py` holds this seam.

## Prove it in a browser

`deckcheck/scripts/review_proof.mjs` drives the app in headless Chrome. It needs Node 24 and Chrome, and `CHROME` sets the browser path. It starts its own server on a free port and stops it at the end, so stop any other review server first.

```bash
node deckcheck/scripts/review_proof.mjs artifacts/review-proof
```

It processes the insurance, solar, and fmcg golden meetings from the home screen, decides every change through the page, applies, and saves screenshots. For each meeting it runs `changeset apply` on a copy of the committed ChangeSet with the same decisions and checks that both decks have the same bytes.

It then runs one private meeting through the real `--maker` command. It creates `private/meetings/my meeting/` with a copy of the solar deck as `before.pptx` and a short made-up `notes.md`, and puts a fake `claude` first on the server's `PATH`. It checks that the meeting is listed, that cancelling the consent confirm sends nothing, and that confirming processes it. It checks that a Keep click the engine refuses shows its message, and that Apply matches `changeset apply`. It sends traversal ids and checks each answers 404. Then it deletes the meeting and its review. If `private/meetings/my meeting/` already exists, the proof stops before it starts. It prints `REVIEW PROOF PASS` or `REVIEW PROOF FAIL` and exits 0 or 1. It starts those three meetings over, so their decisions under `artifacts/review/` are lost.
