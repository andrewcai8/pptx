# The edit engine

The edit engine is the first piece of meeting-to-slides. A maker reads a meeting and writes a ChangeSet, a JSON file of edits to one source deck. The engine checks the ChangeSet against the deck, writes a copy with every change applied for review, and then writes the final deck from the reviewer's decisions. It never writes to the source deck.

These commands print slide text, so the confidentiality rules in `CLAUDE.md` apply to every deck you run them on.

## Commands

Run every command from the repo root. A ChangeSet names its source deck by a path relative to the repo root.

```bash
uv run --project deckcheck changeset validate private/meeting/changeset.json
uv run --project deckcheck changeset execute private/meeting/changeset.json --out private/meeting/executed.pptx --review artifacts/meeting/review.json
uv run --project deckcheck changeset apply private/meeting/changeset.json --out private/meeting/final.pptx
```

- `validate` checks every change against the source deck and prints one line per change with the text it replaces.
- `execute` writes the deck with every change as the maker wrote it. `--review` also writes the review view, described below.
- `apply` writes the final deck from a fresh copy of the source. It replays the changes marked `keep_new` or `edited` and leaves out the ones marked `keep_old`. It refuses to run while any decision is pending.

| exit | meaning |
|---|---|
| 0 | done |
| 1 | the ChangeSet has problems, listed one per line as `<change id> <field>: <message>`, or `apply` found pending decisions. Nothing is written. |
| 2 | the ChangeSet file cannot be read, the source file is not a deck, or an output path (`--out` or `--review`) is the source deck, the ChangeSet, a folder, or the other output |

## The ChangeSet

`deckcheck/src/deckcheck/changeset/changeset.schema.json` is the full format. A ChangeSet holds:

- `meeting`, with a `title` and a `date`.
- `source`, the deck's `path` and `sha256`. A deck whose hash differs is refused.
- `asks`, what the meeting asked for. Each has an `id`, a `text`, and `refs`.
- `changes`, each with an `id`, the `ask_id` it serves, a one-sentence `rationale`, `refs`, an `op`, and a `decision`.
- `flags`, questions the maker could not settle, and `held`, things discussed and deliberately left alone. Each names its slides and refs. Neither touches the deck.

A ref is `{"t": "00:05:46", "speaker": "Grace Adeyemi", "quote": "..."}`, with the quote copied verbatim from the transcript.

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

Paragraph indexes count every paragraph in the shape, blank ones included. Text reads the way python-pptx's `paragraph.text` reads it, with a line break as `\v`. A slide placed after another follows it wherever that slide ends up. Several slides placed after one slide follow it in ChangeSet order.

The engine reads each change's old text from the source deck and decides from the op whether it is text-only or structural. A maker never writes either. A field such as `before` is refused.

### Decisions

The review app writes each change's decision into the ChangeSet: `"keep_new"`, `"keep_old"`, or `{"edited": "<text>"}`. A missing decision is `"pending"`. `edited` replaces what the change writes: `new` for `replace_text` and `set_cell`, `text` for `insert_paragraph`, a number for `set_chart_value`, and one paragraph per line for `fill_placeholder`. `add_slide`, `delete_slide`, and `move_slide` take `keep_new` or `keep_old` only. A `fill_placeholder` whose `add_slide` is kept old is dropped, and `apply` says so.

The engine checks every change against every other whatever the decisions, so any set of decisions gives a valid deck. All `keep_old` gives the source deck byte for byte. All `keep_new` gives the executed deck byte for byte. The same decisions always give the same bytes.

## The review view

`execute --review` writes a JSON file in the shape of `deckcheck/src/deckcheck/changeset/review.schema.json`. It lists every slide with its source and executed positions, and every change with its shape, the text before and after, the quoted span, what it depends on, and the decisions it admits.

## Known limits

- A chart whose workbook is `.xlsb` gets a new `.xlsx` workbook that holds the old workbook's values only. Formulas and formatting in that workbook are lost. The solar deck's six charts are the only `.xlsb` charts in the corpus.
- There is no `outline` command yet that lists the slide, shape, and paragraph ids a maker can address. Until there is, a bad id's error message lists the ids that exist.
- `replace_text` cannot add or remove a line break, and no op changes fonts, sizes, colors, or positions.
