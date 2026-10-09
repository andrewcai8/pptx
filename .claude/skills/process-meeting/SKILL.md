---
name: process-meeting
description: Turn a meeting transcript or notes into a ChangeSet, the JSON list of edits to one source deck that the edit engine checks and applies for a consultant to review. Use when asked to process a meeting, turn a transcript into deck edits, or prepare slide changes from a call, and when `meeting process` runs this skill headless.
---

# Process a meeting

You read a meeting and the deck it discussed, and you write one ChangeSet: what the meeting asked for, the edits that serve each ask, the asks too vague to act on, and the things discussed that must stay as they are. The edit engine applies the edits. A consultant then reviews each one.

You never edit the deck. You never write a `.pptx`, open the deck with python-pptx, or change any file outside the working folder. The ChangeSet is your only product.

## Where you run

`meeting process` runs you in a working folder it staged for this run, `<out>/stage/`. Everything you need is in it:

- `transcript.md`, one turn per line as `[00:41:07] Name (Role, Org): text`, or `notes.md` as the recorder exported it
- `data/*.csv`, the client files that new numbers come from, when the meeting has any
- `before.pptx`, the source deck
- this skill, and `docs/changeset.md`, which describes every op and field

Your file tools reach nothing outside this folder, and nothing outside it is yours to read. Run every command from it.

## Confidentiality

A client deck and its meeting are confidential, and everything you read leaves the machine.

- When `meeting process --shareable` started you, the user already answered the ask-first question in the repo's `CLAUDE.md`. Its prompt says so. Do not ask again.
- If you were started any other way, do not run these steps yourself. Ask the user whether the deck and the meeting may be shared with an AI service under their firm's policy. If they say yes, run `uv run --project deckcheck meeting process <meeting folder> --out <a folder under private/ or artifacts/> --shareable` from the repo root. If the answer is no or unsure, stop.
- Never commit anything from the working folder.

## Tools

Use only these:

- Read to read a file, Grep to search one, and Glob to list files. Read a long file one part at a time.
- Write to create or replace a file. Write only inside the working folder.
- These four commands, each exactly as written, with nothing added before or after:

```bash
uv run --project deckcheck changeset outline before.pptx --json outline.json
uv run --project deckcheck changeset validate changeset.json
uv run --project deckcheck changeset execute changeset.json --out executed.pptx --review review.json
uv run --project deckcheck deckcheck check executed.pptx --out check
```

Do not read or write files with shell commands such as `cat`, `sed`, `grep`, `python3`, or a `>` redirect, and do not set shell variables. A headless run denies every other command.

## Steps

### 1. Read the meeting and outline the deck

Read the whole meeting and every data file. Then list the deck:

```bash
uv run --project deckcheck changeset outline before.pptx --json outline.json
```

The command prints one line with the slide count and the deck's sha256, and writes the outline to `outline.json`. Read `outline.json` with Read, in parts, and search it with Grep. It holds:

- `deck`, with the `path` and `sha256` the ChangeSet's `source` copies
- `layouts`, each with a `name` and the `placeholders` an added slide can fill, by `id`
- `slides`, each with its `index`, `id`, `layout`, `title`, and `shapes`. Each shape has an `id`, a `name`, and a `kind`. A text shape lists its `paragraphs` in order, so the first is paragraph 0. A table lists its `rows`, each a list of cell texts, with `null` for a merged cell. A chart lists its `series`, each with an `index`, a `name`, and `points` that hold a `point` number, a `category`, and a `value`.

Every text is a JSON string, read the way the engine reads it, so a value copied from `outline.json` can go straight into `old`.

People say slide numbers. "Slide five" is the slide whose `index` is 5. The ChangeSet names slides by `id`, so look each one up.

### 2. Build the ask list

Go through the meeting turn by turn. Sort every topic that touches the deck or its content into exactly one of these:

- **An ask.** Someone asked for a change and the meeting settled what it is. It becomes an entry in `asks` and one or more `changes`.
- **Held.** The meeting discussed something and decided it stays as it is. It becomes an entry in `held`.
- **A flag.** Someone asked for a change but the meeting never settled which slide, which value, or which option. It becomes an entry in `flags`.

These rules decide the hard cases:

- The last word wins. When a speaker corrects a number, changes their mind, or drops a suggestion, use the final answer. Record the abandoned answer as held, and keep it out of the whole deck, speaker notes included.
- "Leave it", "don't touch it", "not this round", "park it", and "we'll revisit next week" are held. Name the slides they are about.
- A suggestion that someone rejected, a follow-up that happens outside the deck, a logistics item, a topic someone said to keep off the slide, and any name or number the meeting said must not appear are held. These are the things a careless edit would add, so write each one down.
- An ask is vague when the meeting left a real choice open: two slides it could mean, two values it could take, or "make it better" with no direction. Never guess. Do not edit the candidate slides for it. Write a flag that asks the concrete question, lists the options the meeting mentioned, names every candidate slide in `slides`, and cites every turn where the ask was made or questioned.
- Small talk with no bearing on the deck needs no entry.

### 3. Write each change

Each change is one op on one place in the deck. Copy every `old` value from the outline character for character. Paste the JSON string literal.

- Change the smallest span that carries the edit. To update a figure, replace `ca. 25%` and leave the rest of the sentence alone. To reword a bullet, replace that bullet's text in its own paragraph. The review keeps every other clause on the slide, so a rewritten sentence that drops a clause fails.
- `old` must occur exactly once in the shape. If it repeats across paragraphs, add `paragraph` with the index. If it repeats inside one paragraph, quote more of the text around it.
- Find every place the changed fact appears. Grep `outline.json` for each spelling of the old value, such as `120`, `$120m`, `c.$120m`, and `+12%`. Check titles, headers, body text, table cells, chart labels, chart points, and footnotes on each slide the meeting named, and on any other slide where the meeting said the figure appears. A figure drawn as a chart bar needs `set_chart_value` as well as the label next to it, or the bar and the label disagree.
- Edit only the slides the meeting asked to change. A slide the meeting did not ask about keeps every byte.
- To delete a slide, use `delete_slide`. To move one, use `move_slide` with `after` set to the id of the slide it should follow, or null to make it first. "Put the summary first" is `{"slide": <id of the summary>, "after": null}`.
- To add a slide, use `add_slide` with a layout name from the outline's layouts. When the meeting says "same look as slide N", use slide N's layout. Then write one `fill_placeholder` per placeholder you fill, using the placeholder ids listed under that layout: the title, and the body as one entry per bullet.
- When the meeting gives the wording, use it as said. When it gives the content but not the words, write it in the deck's voice and use only facts from the meeting, a data file, or the slide itself.
- Add only what the meeting asked for. Add no caveats, no extra bullets, no follow-ups, and no speaker notes. Never write a question, a placeholder, or "to be confirmed" into the deck. Questions go in `flags`.
- When the meeting asks to cite a source, or a new figure comes from a data file and the slide already has a footnote or source line, add one line that names the source. Use `insert_paragraph` after the last footnote and follow the existing numbering.

New text must meet the house style, which `deckcheck check` enforces. A bullet must not end in `.`, `;`, or `,`, so end no paragraph you write with one. A title stays under 150 characters. Never write `TBD`, `XX`, `[insert`, or `???`.

### 4. Take numbers only from the meeting or a data file

Every number you write must be stated in a transcript turn, held in a data file, or already on the slide. Never compute, round, or estimate a new one. When the meeting and a data file disagree, follow what the meeting decided last.

Cite each number. Put the turn that states it in the change's `refs`. When it comes from a data file, the `rationale` names the file, the row, and the column, as in `data/sales.csv row revenue_2025, column value`.

### 5. Write the ChangeSet and validate it

Write `changeset.json` with Write, in this shape. `docs/changeset.md` describes every op and field.

```json
{
  "meeting": {"title": "<the meeting header's title>", "date": "YYYY-MM-DD"},
  "source": {"path": "before.pptx", "sha256": "<sha256 from the outline>"},
  "asks": [{"id": "new-price", "text": "What was asked, in one sentence.", "refs": [REF]}],
  "changes": [
    {"id": "title-price", "ask_id": "new-price", "rationale": "One sentence.", "refs": [REF],
     "op": {"kind": "replace_text", "slide": 260, "shape": 4, "old": "ca. 25%", "new": "ca. 28%"}}
  ],
  "flags": [{"id": "which-chart", "question": "The concrete question, with the options.", "slides": [262, 265], "refs": [REF]}],
  "held": [{"id": "logo-size", "text": "What stays as it is, and why.", "slides": [256], "refs": [REF]}]
}
```

A `REF` is `{"t": "00:03:12", "speaker": "Ana Ruiz", "quote": "Use 28, that's the new figure."}`. The `t` is the turn's timestamp, the `speaker` is the name before the parenthesis, and the `quote` is copied verbatim from that turn's text. Every ask, change, flag, and held entry needs at least one ref. A `rationale` is one sentence on one line. Ids use letters, digits, `.`, `_`, and `-`. Leave out `decision`, which the review app writes.

The ops are `replace_text` (`slide`, `shape`, `old`, `new`, optional `paragraph`), `insert_paragraph` (`slide`, `shape`, `after`, `text`), `set_cell` (`slide`, `shape`, `row`, `col`, `old`, `new`), `set_chart_value` (`slide`, `shape`, `series`, `point`, `old`, `new` as numbers), `add_slide` (`layout`, `after`), `fill_placeholder` (`slide` as the add_slide change's id, `shape` as the placeholder id, `paragraphs` as `[{"text": ..., "level": 0}]`), `delete_slide` (`slide`), and `move_slide` (`slide`, `after`).

Then validate:

```bash
uv run --project deckcheck changeset validate changeset.json
```

Exit 0 prints `VALID` and one line per change with the text it replaces. Exit 1 prints `INVALID` and one `<change id> <field>: <message>` line per problem. Fix every problem in the ChangeSet and validate again until it prints `VALID`. Each message names what exists, such as the deck's slide ids or a shape's text, so read it before you change anything.

### 6. Execute and check the house style

```bash
uv run --project deckcheck changeset execute changeset.json --out executed.pptx --review review.json
uv run --project deckcheck deckcheck check executed.pptx --out check
```

`execute` writes the edited copy and `review.json`, which lists each change's text before and after. Read `review.json` and the after text of every change. It must say what the meeting decided and read as a finished slide.

`check` exits 1 whenever the deck has any violation, and most source decks already have some. A violation is yours when it names a slide you changed or added and the text it quotes is text you wrote. Fix each one by writing the ChangeSet again, then validate, execute, and check again. Leave the source deck's own violations alone.

## Done

You are done when `validate` prints `VALID`, `execute` has written the deck, and `check` reports no violation in text you wrote. Reply with the asks, the changes, the flags, and the held items, one line each, and the path of the ChangeSet.
