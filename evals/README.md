# Golden scenarios

Each scenario is a recorded meeting about a real deck, plus the deck changes the meeting asked for. `score.py` decides by script whether an output deck made those changes and nothing else. `prove.py` builds known-good and known-bad outputs for every scenario and checks that the scorer passes the good one and fails each bad one for the declared reasons.

A scenario directory holds these files:

- `transcript.md` is the meeting. It has a short header, then one turn per line in the form `[00:04:10] Name (Role, Org): text`.
- `expected.yaml` lists the changes, the things discussed that must not become edits, and the facts the output must and must not contain.
- `build.py` declares the outputs as `@variant` functions. At least one passes, such as the edit the meeting settled on and faithful rewordings of it. At least two fail, each with its exact set of `(code, slide)` failures.
- `data/*.csv` holds synthetic client data that a number can cite.

Source decks are never committed. A public scenario names a deck in `.claude/skills/verify-pptx/corpus/known-good.yaml` by id and sha256, and the scorer reads it from the corpus cache. Generated decks and `score.json` files go under `artifacts/evals/`.

Run the scorer's unit tests with `uv run --project deckcheck pytest evals`.

## The one contract a maker keeps

The maker edits a copy of the source deck. The scorer tells the slides apart by their slide ids (`p:sldId/@id`), which python-pptx and PowerPoint keep when a slide is edited, deleted, or moved. A new slide gets a fresh id. An output that shares no slide id with the source fails with one `structure` failure, because the scorer cannot tell which slide is which.

## Run the proof

Run every command from the repo root.

```
uv run --project deckcheck python evals/prove.py
uv run --project deckcheck python evals/prove.py insurance-workshop-prep
```

`prove.py` proves only the public scenarios in `evals/`. With no names, it proves all of them and then checks the composition of the set. The set needs at least 5 scenarios on at least 4 decks, and its negatives must cover every failure code. With names, it proves only those scenarios and skips the set checks. It prints one `ok` or `BAD` line per variant, then `PROOF PASS` or `PROOF FAIL`, then the evidence directory. A build script that raises is a `BAD` line and exit 2.

`prove.py` rebuilds every variant from a fresh copy of the source deck on each run, so the build scripts need no separate command. The outputs and their `score.json` files land under the evidence directory as `<scenario>/<variant>/`.

## Score one output

```
uv run --project deckcheck python evals/score.py insurance-workshop-prep path/to/output.pptx [--out DIR]
```

The scenario argument is a directory, or a name looked up in `evals/` and then in `$GOLDEN_PRIVATE_DIR`. The command prints one line, `SCENARIO PASS <name> (<n> checks, <k> intent checks deferred)` or `SCENARIO FAIL: [<code>] <message>; ...`. It writes `score.json` to `--out`, or else to the output's directory.

| exit | meaning |
|---|---|
| 0 | pass |
| 1 | fail |
| 2 | bad scenario, bad arguments, or an unreadable output |
| 3 | the source deck is unreachable |

## Write expected.yaml

Slide numbers are always source slide numbers, the ones in the deck's `outline.md`. An added slide is named by its change id. The scorer derives the output slide order from the change list, so nobody writes output positions.

```yaml
deck:
  corpus: insurance-regulator-2017
  sha256: 5dcf6300cfc6c95ab5f9ba184519835380fb60b30b0dea20a47eeb7739543352
changes:
  - id: c1
    kind: update-number            # edit-text, update-number, or restyle
    intent: One sentence.
    said: ["00:04:10", "00:11:52"]
    slides:
      5:
        require:
          - {percent: "48%", where: title, from: {data: data/client-survey.csv, row: millennials-2027, column: share_pct}}
        forbid:
          - {percent: "50%", where: title}
          - {percent: "45%", superseded: "00:04:10"}
    intent_checks: [What only a reader can judge.]
  - {id: c2, kind: delete-slide, slide: 2, intent: ..., said: [...]}
  - {id: c3, kind: move-slide, slide: 14, after: 11, intent: ..., said: [...]}
  - {id: c4, kind: add-slide, after: 13, layout: "D. Title only", intent: ..., said: [...], require: [...], forbid: [...]}
non_changes:
  - {id: n1, kind: not-a-change, said: [...], why: ..., slides: [1], absent: [{text: "Oversight Office"}]}
  - {id: n2, kind: ambiguous, said: [...], why: ..., slides: [10, 16, 18], flag: The question the maker should ask.}
```

The loader lints every field against the real deck and transcript, and a bad scenario exits 2 with the field named. These are the main rules:

- Each `said` and `superseded` timestamp is a transcript turn.
- A `require` with a number needs `from`. The value is `{said: <ts>}` when the turn states it, `{data: <csv>, row: <first-column key>, column: <header>}` when the cell value appears in the fact, or `{slide: <n>}` when it is already on that source slide.
- A `require` on an edited slide must not already hold on the source slide, or it could not show the edit happened.
- A plain `forbid` is an old value, so it must be on the source slide. A `superseded` forbid is the abandoned answer from a change of mind. Its timestamp is the turn where that answer was said, so that turn must state it. It must be absent from the source slide.
- Every `absent` fact must be absent from the whole source deck.
- An added slide's `layout` must be a layout that some source slide uses.
- A non-change cannot name a slide that a change edits or deletes. It can name a moved slide, because a moved slide keeps its content.

A fact names the thing that must be true, not one phrasing of it. Each fact has exactly one of these keys:

| key | value | matches |
|---|---|---|
| `money` | `"$100M"` | any amount of the same currency and value: `$100M`, `$100 million`, `$100m`, `US$100M`, `USD 100 million`, `$0.1bn` |
| `percent` | `"48%"` | `48%`, `48 %`, `48 per cent`, `48 percent` |
| `count` | `"6 weeks"` | the number in digits or words up to twenty, then the unit, with at most one word between: `6 weeks`, `six-week`, `6 calendar weeks` |
| `chart` | `410` | a value in the slide's chart data, as python-pptx reads it from the chart part. It takes no `where` |
| `text` | a string or a list | the load-bearing concept, with the few wordings it needs, such as `["cap", "limit"]` |

A `require` passes when the fact is on the slide. A `forbid` or `absent` fails when it is there. Matching ignores case, Unicode composition, runs of whitespace, and curly quotes. A match cannot sit inside a longer word or number, so `39%` matches `+39%` but not `139%` or `1.39%`, and `day` does not match `days`. `evals/test_facts.py` lists the cases.

Anything that is really about wording, such as tone, which phrase was used, or where a bullet sits, belongs in `intent_checks`.

## What the script decides and what it defers

The script decides these checks, in this order:

1. `source` checks that the source deck still has its pinned hash and that the output is not the source file.
2. `structure` maps every output slide to a source slide by slide id, or to an added slide in order of appearance. It reports a kept slide that is missing, a deleted slide that is still there, a new slide that no change asks for, an added slide that is missing, and a slide out of order. When two slides could explain one displacement, it blames the slide a change moved or added.
3. `scope` runs on every slide the mapping pairs, even when `structure` failed. A slide no change edits must look the same: the same text, the same slide XML, and the same related parts (charts and their embedded workbooks, images and other media, notes), followed recursively. A notes page with no text counts as no notes page. A changed slide gets `guessed` when an ambiguous non-change names it, `non-change` when a not-a-change names it, and `scope` otherwise. On an edited slide, every source paragraph must still be there, unless it holds one of that slide's plain `forbid` values or starts with the house-style source prefix. An edit target that did not change fails `missing`. A change to the deck's slide layouts, masters, or themes is one `scope` failure.
4. `missing` and `forbidden` check the `require` and `forbid` facts on each changed slide, and the `absent` facts across the deck.
5. `layout` checks that each added slide uses its declared layout.
6. `style` reports house-style violations that the source deck did not already have, matched through the slide mapping.

Some python-pptx getters add XML when code only reads a deck. The XML comparison ignores exactly the ones that render the same as no element. `READ_ARTIFACTS` in `scenario.py` is that table, and `evals/read_artifacts.py` measures it over the corpus:

| getter | adds | scored |
|---|---|---|
| `paragraph.level`, `paragraph.alignment` | an empty `a:pPr` in `a:p` | ignored |
| `paragraph.font` | an empty `a:defRPr` in `a:pPr` | ignored |
| `run.font`, `run.hyperlink` | an empty `a:rPr` in `a:r` | ignored |
| `shape.line.fill` | an empty `a:ln` in `p:spPr` | ignored |
| `slide.notes_slide` | a notes page with no text | ignored |
| `run.font.color` | an empty `a:solidFill` in `a:rPr` | a change |
| `shape.line.color` | `a:ln` with an empty `a:solidFill` | a change |
| `chart.chart_title` | a `c:title` | a change |

The last three render differently, so a maker must not touch them on a slide nobody asked about. Any other XML difference, even an empty element such as `a:buNone` or `a:noFill`, is a change. A tool that re-serializes the whole slide XML, such as a PowerPoint save or LibreOffice, may change an untouched slide's XML in other ways. Scoring the output of such a tool may need the table extended, from a measurement like `read_artifacts.py`.

The intent checker judges the rest. That covers every `intent_checks` line, such as tone, wording, and placement, and whether the maker raised each ambiguous `flag`. `score.json` lists them under `deferred`. For an ambiguous ask, the script proves only that the maker did not guess.

## Private scenarios

A scenario built on a client deck lives in `$GOLDEN_PRIVATE_DIR/<name>/` and never enters the repo. It names its deck by file and hash:

```yaml
deck: {file: input.pptx, sha256: <sha256 of input.pptx>}
```

The loader rejects a `file:` deck in a scenario inside the repo. Only `score.py` looks in `$GOLDEN_PRIVATE_DIR`. A name found in both `evals/` and `$GOLDEN_PRIVATE_DIR` is an error.
