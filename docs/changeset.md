# The edit engine

The edit engine is the first piece of meeting-to-slides. A maker, the `process-meeting` skill, reads a meeting and writes a ChangeSet, a JSON file of edits to one source deck. The engine checks the ChangeSet against the deck, writes a copy with every change applied for review, and then writes the final deck from the reviewer's decisions. It never writes to the source deck.

These commands print slide text, so the confidentiality rules in `CLAUDE.md` apply to every deck you run them on.

## Commands

Run every command from the repo root. A ChangeSet names its source deck by a path relative to the repo root.

```bash
uv run --project deckcheck changeset outline private/meeting/before.pptx --json private/meeting/outline.json
uv run --project deckcheck changeset validate private/meeting/changeset.json
uv run --project deckcheck changeset execute private/meeting/changeset.json --out private/meeting/executed.pptx --review artifacts/meeting/review.json
uv run --project deckcheck changeset apply private/meeting/changeset.json --out private/meeting/final.pptx
```

- `outline` lists what a ChangeSet can address. It prints the deck's path and sha256, the layouts `add_slide` accepts with their text placeholders, and each slide's index, id, layout, and title. Under each slide it lists every shape by id, name, and kind, with numbered paragraphs (`p0`), table cells (`r1c4`, or `merged`), or chart series and points. Every text is a JSON string literal read the way the engine reads it, so it can be pasted into `old`. `--json` writes the same outline as JSON to a file instead, and prints one line with the slide count and sha256. A long deck's outline runs to tens of thousands of characters, more than an agent's console shows.
- `validate` checks every change against the source deck and prints one line per change with the text it replaces.
- `execute` writes the deck with every change as the maker wrote it. `--review` also writes the review view, described below.
- `apply` writes the final deck from a fresh copy of the source. It replays the changes marked `keep_new` or `edited` and leaves out the ones marked `keep_old`. It refuses to run while any decision is pending.

| exit | meaning |
|---|---|
| 0 | done |
| 1 | the ChangeSet has problems, listed one per line as `<change id> <field>: <message>`, or `apply` found pending decisions. A ChangeSet that is not JSON is listed as `changeset: not JSON`. A source deck that is missing or cannot be read is listed under `source.path`, and one with another hash under `source.sha256`. Nothing is written. |
| 2 | the ChangeSet file is missing, is a folder, cannot be read, or is not UTF-8 text, the deck given to `outline` or named as the source is not a deck, an output path (`--out` or `--review`) is the source deck, the ChangeSet, a folder, or the other output, or `outline --json` is the deck |

## The ChangeSet

`deckcheck/src/deckcheck/changeset/changeset.schema.json` is the full format. A ChangeSet holds:

- `meeting`, with a `title` and a `date`.
- `source`, the deck's `path` and `sha256`. A deck whose hash differs is refused.
- `asks`, what the meeting asked for. Each has an `id`, a `text`, and `refs`.
- `changes`, each with an `id`, the `ask_id` it serves, a one-sentence `rationale`, `refs`, an `op`, and a `decision`.
- `flags`, questions the maker could not settle, and `held`, things discussed and deliberately left alone. Each names its slides and refs. Neither touches the deck.

A ref is `{"t": "00:41:07", "speaker": "Ana Ruiz", "quote": "..."}`, with the quote copied verbatim from the transcript.

A slide is named by its slide id (`p:sldId/@id`) and a shape by its id on that slide (`p:cNvPr/@id`), including shapes inside groups. These ids survive edits, deletions, and moves, so they mean the same thing in every deck the engine writes.

| op | what it does | review |
|---|---|---|
| `replace_text` | replaces the quote `old` with `new` in one shape. `old` must occur exactly once in the shape. When it repeats across paragraphs, `paragraph` names the one. | text-only |
| `insert_paragraph` | adds a paragraph `text` after paragraph `after`, styled like it | text-only |
| `set_cell` | sets table cell `row`, `col` from `old` to `new`, paragraphs joined by `\n` | structural |
| `set_chart_value` | sets point `point` of series `series` from `old` to `new`, in the chart and in its embedded workbook | structural |
| `add_slide` | adds a slide on one of the deck's own layouts, by name, after source slide `after`, or first when `after` is null | structural |
| `fill_placeholder` | fills placeholder `shape` on the slide that the `add_slide` change `slide` adds, one entry per paragraph with an optional `level` | structural |
| `delete_slide` | deletes slide `slide` | structural |
| `move_slide` | moves slide `slide` after source slide `after`, or first when `after` is null | structural |

The text a change writes is `new`, `text`, a fill's paragraphs, or an `edited` decision. It may hold a tab, which the deck keeps as a tab. It may not hold any other control character, such as `\f`, `\r`, or `\x1f`, and a change with one is refused with the character and its index, counted from the start of that text. A line feed and a line break keep their rules. `\n` splits paragraphs in `set_cell` and in an edited `fill_placeholder` and is refused elsewhere, and `\v` is a line break. So a review app turns a pasted `\r\n` into `\n` before it writes a decision.

A chart point's workbook cell must hold a plain number. A blank cell, a formula, text, a true/false value, an error, or a date is refused. Paragraph indexes count every paragraph in the shape, blank ones included. Text reads the way python-pptx's `paragraph.text` reads it, with a line break as `\v`. A slide placed after another follows it wherever that slide ends up. Several slides placed after one slide follow it in ChangeSet order.

The engine reads each change's old text from the source deck and decides from the op whether it is text-only or structural. A maker never writes either. A field such as `before` is refused.

### Decisions

The review app writes each change's decision into the ChangeSet: `"keep_new"`, `"keep_old"`, or `{"edited": "<text>"}`. A missing decision is `"pending"`. `edited` replaces what the change writes: `new` for `replace_text` and `set_cell`, `text` for `insert_paragraph`, a number for `set_chart_value`, written with `.` for decimals and optionally commas between groups of three digits (`1234.5` or `1,234.5`), and one paragraph per line for `fill_placeholder`. The edited text must meet the same rules as the maker's, so an empty `insert_paragraph` is refused. `add_slide`, `delete_slide`, and `move_slide` take `keep_new` or `keep_old` only. A `fill_placeholder` whose `add_slide` is kept old is dropped, and `apply` says so.

The engine checks every change against every other whatever the decisions, so any set of decisions gives a valid deck. All `keep_old` gives the source deck byte for byte. All `keep_new` gives the executed deck byte for byte. The same decisions always give the same bytes.

## The review view

`execute --review` writes a JSON file in the shape of `deckcheck/src/deckcheck/changeset/review.schema.json`. It lists every slide with its source and executed positions, and every change with its shape, the text before and after, the quoted span, what it depends on, and the decisions it admits.

## Run the maker headless

`meeting process` runs the `process-meeting` skill without a person at the keyboard, then validates what it wrote.

```bash
uv run --project deckcheck meeting process private/meetings/q3 --out private/meetings/q3/run1 --shareable
```

`--out` must sit inside `artifacts/` or `private/`. `--shareable` answers the ask-first question in `CLAUDE.md`, and the command refuses to start without it. Before the run it checks `claude auth status --json`.

The run sees only a staged copy of what it needs. The command replaces `<out>/stage/` on every run and copies these into it:

- `transcript.md`, or `notes.md` when the folder has no transcript
- every `data/*.csv`
- the deck, as `before.pptx`. It is `before.pptx` in the meeting folder unless `--deck` names another.
- the skill, at `.claude/skills/process-meeting/SKILL.md`, and this file, at `docs/changeset.md`
- `deckcheck`, a link to the engine, so `uv run --project deckcheck` works there

Nothing else in the meeting folder is staged, so `after.pptx`, `expected.yaml`, `changeset.json`, `build.py`, and other meetings stay out of reach. The command then runs this from `<out>/stage/`:

```bash
claude -p '<prompt>' --tools Read,Write,Grep,Glob,Bash --restricted --strict-mcp-config --permission-mode dontAsk \
  --allowedTools 'Edit(./**)' \
    'Bash(uv run --project deckcheck changeset outline before.pptx --json outline.json)' \
    'Bash(uv run --project deckcheck changeset validate changeset.json)' \
    'Bash(uv run --project deckcheck changeset execute changeset.json --out executed.pptx --review review.json)' \
    'Bash(uv run --project deckcheck deckcheck check executed.pptx --out check)' \
  --output-format json --no-session-persistence
```

- `--tools` gives Claude five built-in tools and no tool that reaches the network.
- `--restricted` keeps Read, Write, Grep, and Glob inside `<out>/stage/`, and a file reached through a link that leads out of it counts as outside. It also ignores the user, project, and local settings files, so their permission rules and hooks do not load. Managed settings, which an organization's admin sets, still apply.
- `--strict-mcp-config` loads no MCP server, because the command passes no `--mcp-config`.
- `--permission-mode dontAsk` denies every call that no rule allows, so a headless run never waits on a prompt. `Edit(./**)` lets Write create files in `<out>/stage/`. The `Bash` rules allow the four commands as written, which the prompt and the skill name. Claude Code also accepts them followed by `2>&1` or by a redirect into a file in `<out>/stage/`, and it runs plain read-only commands such as `ls` on files there without a rule. Any other command, a variable, or a path outside `<out>/stage/` is denied.
- The commands name only files inside `<out>/stage/`, so a meeting or `--out` path with spaces does not change them.

Afterwards the command writes Claude's JSON result to `<out>/claude.json` and copies `<out>/stage/changeset.json` to `<out>/changeset.json`, naming the deck by its path from the repo root, as every other command expects. `outline.json`, `executed.pptx`, `review.json`, and `check/` stay in `<out>/stage/`.

| exit | meaning |
|---|---|
| 0 | the ChangeSet is valid, and the command prints what `changeset validate` prints |
| 1 | Claude wrote no ChangeSet, or it is invalid or names another deck than `before.pptx` |
| 2 | bad input, such as a missing transcript, a missing deck, `--out` outside `artifacts/` and `private/`, a meeting or deck inside `<out>/stage/`, or no `--shareable` |
| 3 | the `claude` CLI is missing or not logged in, and the command prints ``claude CLI not found or not logged in; run `claude` once to log in`` |

These flags ran through Claude Code 2.1.293 against a local stand-in for the API. It allowed the four commands and denied each read, write, and command that reached outside `<out>/stage/`, including one that a user settings rule allowed. `deckcheck/tests/test_meeting.py` pins the argv and the staged files. No live `claude -p` run against Claude has tested them.

## Known gaps

- A chart whose workbook is `.xlsb` gets a new `.xlsx` workbook that holds the old workbook's values only. Formulas, defined names, cell styles, and date formats in that workbook are lost, and a text cell that begins with `=` becomes a formula. The solar deck's six charts are the only `.xlsb` charts in the corpus.
- `delete_slide` leaves a slide's part in the package when another slide links to it, for example through a click action that jumps to it. The part drops out of the slide list, and the link still points at it. No corpus deck has such a link, and it is not known whether PowerPoint asks to repair the file. Agenda decks often have them.
- `replace_text` cannot add or remove a line break, and no op changes fonts, sizes, colors, or positions.
- A `replace_text` that rewrites text across differently formatted runs puts the new text in the first run it touches. So rewriting a bold lead-in can leave new plain words bold. `Luxury casualwear: keeps increasing` rewritten to `Luxury streetwear: keeps growing` shows bold up to "keeps grow". The review app shows every change, so the consultant can fix it with Edit myself.
- Some decks hold a line feed inside a paragraph's text. `replace_text` cannot quote across one, and `set_cell` refuses a cell that holds one, because `\n` there would not say where the cell's paragraphs split.
