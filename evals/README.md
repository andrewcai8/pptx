# Golden scenarios

Each scenario is a recorded meeting about a real deck, plus the deck changes the meeting asked for. `score.py` decides by script only what a script can decide objectively. It checks that the right slides changed, were added, deleted, or moved, that every other slide keeps its slide XML and related parts, that an edited slide kept its on-slide text, its pictures, charts, tables, embedded objects, and groups, and its chart values, except what the edit targets, and that the required facts are there and the abandoned ones are not. What an edited or added slide gains is left to the intent checker, through each scenario's `intent_checks`. `prove.py` builds known-good and known-bad outputs for every scenario and checks that the scorer passes the good one and fails each bad one for the declared reasons.

A scenario directory holds these files:

- `transcript.md` is the meeting. It has a short header, then one turn per line in the form `[00:04:10] Name (Role, Org): text`.
- `expected.yaml` lists the changes, the things discussed that must not become edits, and the facts the output must and must not contain.
- `build.py` declares the outputs as `@variant` functions. At least one passes, such as the edit the meeting settled on and faithful rewordings of it. At least two fail, each with its exact set of `(code, slide)` failures. A variant declared with `intent=(id, phrase)` is a wrong deck whose only fault is something it added. The script passes it, and `prove.py` checks that it passes and that the change or non-change `id` has an intent check containing `phrase`, so every wrong deck the script lets through names the check that catches it.
- `data/*.csv` holds synthetic client data that a number can cite. Only CSV is read. A `from.data` that points at a spreadsheet such as an `.xlsx` exits 2 with `data/<file>: only CSV is read; export the sheet to CSV`.

Source decks are never committed. A public scenario names a deck in `.claude/skills/verify-pptx/corpus/known-good.yaml` by id and sha256, and the scorer reads it from the corpus cache. Generated decks and `score.json` files go under `artifacts/evals/`.

Run the scorer's unit tests with `uv run --project deckcheck pytest evals`.

## The one contract a maker keeps

The maker edits a copy of the source deck. The scorer tells the slides apart by their slide ids (`p:sldId/@id`), which python-pptx and PowerPoint keep when a slide is edited, deleted, or moved. A new slide gets a fresh id. An output that shares no slide id with the source fails with one `structure` failure, because the scorer cannot tell which slide is which.

## How a maker raises a flag

Some asks are ambiguous, and the right output asks a question instead of guessing. The maker writes its questions to `flags.json` in the same directory as the output deck, never into the deck. A speaker note, comment, or slide text that asks the question is an edit like any other, so on a slide an ambiguous ask names it fails `guessed` (see `flag_in_speaker_note` in the insurance scenario).

```json
[
  {"question": "Which regulator slide should be punchier: 10, 16, or 18?", "said": ["00:08:05", "00:08:38"], "slides": [10, 16, 18]}
]
```

Each flag has a `question` and at least one of `said`, the transcript turns it is about, and `slides`, the source slide numbers it is about. A flag raises an ambiguous non-change when it cites one of that non-change's `said` turns or names one of its `slides` or `flag_slides`. A flag that raises none is `unmatched`.

`score.py` reads `flags.json` when it is there. `score.json` lists under `flags` which ambiguous non-changes were raised, which are missing, the unmatched questions, and anything it could not read, and the command prints them on a second line. This is reported, not scored. A missing flag does not fail the deck and an unmatched one does not either, until the intent checker exists to judge whether each question is the right one.

The scorer reads a flags file the way a maker is likely to write it. It reads UTF-8 with or without a byte order mark and ignores keys other than `question`, `said`, and `slides`, such as `id` or `reason`. A turn may drop its hour or leading zeros (`8:05` is `00:08:05`), a slide number may be a numeric string (`"18"`), and a single turn or slide may stand without a list. A file or a flag it still cannot read is reported on the flags line as `unreadable (<reason>)`, and the deck is scored as usual. That includes contrived input, such as a slide number 5000 digits long or JSON nested 100,000 levels deep. The exit code comes from the deck verdict alone.

Any flag that names a slide an ambiguous ask lists raises that ask, whatever the question says. So any flag naming solar slide 10 raises n1, the tariff question. That is loose on purpose, because flags are reported and not scored. Whether the question is the right one is for the intent checker.

## Run the proof

Run every command from the repo root.

```
uv run --project deckcheck python evals/prove.py
uv run --project deckcheck python evals/prove.py insurance-workshop-prep
```

`prove.py` proves only the public scenarios in `evals/`. With no names, it proves all of them and then checks the composition of the set. The set needs at least 5 scenarios on at least 4 decks, and its negatives must cover every failure code except `source` and `unreadable`, which no built public output can earn. `evals/test_scenario.py` covers `unreadable` on a private scenario whose source chart is drawn in 3D. With names, it proves only those scenarios and skips the set checks. It prints one `ok` or `BAD` line per variant, with the covering intent check after a variant left to one, then `PROOF PASS` or `PROOF FAIL`, then the evidence directory. A build script that raises is a `BAD` line and exit 2.

`prove.py` rebuilds every variant from a fresh copy of the source deck on each run, so the build scripts need no separate command. The outputs and their `score.json` files land under the evidence directory as `<scenario>/<variant>/`.

## Score one output

```
uv run --project deckcheck python evals/score.py insurance-workshop-prep path/to/output.pptx [--out DIR]
```

The scenario argument is a directory, or a name looked up in `evals/` and then in `$GOLDEN_PRIVATE_DIR`. The command prints one line, `SCENARIO PASS <name> (<n> checks, <k> intent checks deferred)`, `SCENARIO FAIL: [<code>] <message>; ...`, or `SCENARIO UNREADABLE: [unreadable] <message>; ...`. It writes `score.json` to `--out`, or else to the output's directory.

| exit | meaning |
|---|---|
| 0 | pass |
| 1 | fail |
| 2 | bad scenario, bad arguments, or an unreadable output, such as a file that is not a deck or a chart that points at a part the package lacks, including `SCENARIO UNREADABLE` |
| 3 | the source deck is unreachable |

`SCENARIO UNREADABLE` means every failure is `unreadable`: a check the script could not decide, so a person checks it by hand. It exits 2, not 1, so an unreadable deck never counts as a maker's failure. When a deck also fails a decided check, it is `SCENARIO FAIL` and exits 1, and the `unreadable` entries are listed with the rest.

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
          - {percent: "50%"}
          - {percent: "45%", superseded: "00:04:10"}
    intent_checks: [What only a reader can judge, including what slide 5 may gain and what it must not.]
  - {id: c2, kind: delete-slide, slide: 2, intent: ..., said: [...]}
  - {id: c3, kind: move-slide, slide: 14, after: 11, intent: ..., said: [...]}
  - {id: c4, kind: add-slide, after: 13, layout: "D. Title only", intent: ..., said: [...], require: [...], forbid: [...]}
non_changes:
  - {id: n1, kind: not-a-change, said: [...], why: ..., slides: [1], absent: [{text: "Oversight Office"}], intent_checks: [...]}
  - {id: n2, kind: ambiguous, said: [...], why: ..., slides: [10, 16, 18], flag: The question the maker should ask.}
```

The loader lints every field against the real deck and transcript, and a bad scenario exits 2 with the field named. These are the main rules:

- Each `said` and `superseded` timestamp is a transcript turn.
- A `require` with a number needs `from`. The value is `{said: <ts>}` when the turn states it, `{data: <csv>, row: <first-column key>, column: <header>}`, which must name a CSV file, when the cell value appears in the fact, or `{slide: <n>}` when it is already on that source slide.
- A `require` on an edited slide must not already hold on the source slide, or it could not show the edit happened.
- A plain `forbid` is an old value, so it must be on the source slide. It also names the text the edit may rewrite. A source clause that states it is exempt from preservation (see `lost` below). A `superseded` forbid is the abandoned answer from a change of mind. Its timestamp is the turn where that answer was said, so that turn must state it. It must be absent from the source slide.
- A forbid with `where: deck` is checked on every output slide and its speaker notes, not only the slide it is listed under. Use it where the transcript rules a value out of the whole deck, such as "not tripled anywhere" in retail or "take the eight out completely" in fmcg. Notes count because a deck that goes out as a pptx carries them, as retail's does at 00:07:05. It must be absent from every source slide no change edits, so `3x` on retail slide 5 stays a slide 15 forbid only. It must also be absent from the speaker notes of every source slide the output keeps, edited slides included, because the edit does not ask the maker to rewrite notes.
- Every edit and every added slide needs at least one `intent_checks` line that says what the slide may gain and what it must not, because the script does not judge additions. A delete, a move, or a non-change may carry them too.
- Every `absent` fact must be absent from the whole source deck.
- An added slide's `layout` must be a layout that some source slide uses.
- A non-change cannot name a slide that a change edits or deletes. It can name a moved slide, because a moved slide keeps its content.
- An `ambiguous` non-change's `flag` is the question the maker should put in `flags.json`. Its `said`, `slides`, and `flag_slides` are what a maker's flag must cite to count as raising it, so list every turn where the ask was made. `flag_slides` names the slides a change edits or deletes that the question is about, such as solar slide 10, which holds the $0.08 tariff in doubt. A change there is judged by that change, not as a guess, so a non-change cannot list it under `slides`. Every `flag_slides` slide must be one a change edits or deletes.

A fact names the thing that must be true, not one phrasing of it. Each fact has exactly one of these keys:

| key | value | matches |
|---|---|---|
| `money` | `"$100M"` | any amount of the same currency and value, with the currency before or after the number: `$100M`, `$100MM`, `$100 million`, `$100-million`, `US$ 100 million`, `USD 100 million`, `$0.1bn`, `100m$`, `100 million dollars`, and either end of a range such as `$100-120m` or `$80m to $100m` |
| `percent` | `"48%"` | `48%`, `+48%`, `48 %`, `48 per cent`, `48 percent`, and either end of a range such as `45-48%`, `45–48%`, or `45 to 48 per cent`. `-48%` is a different value |
| `count` | `"6 weeks"` | the number in digits or words up to twenty, then the unit, with up to three words between that are neither a number nor a plural: `6 weeks`, `six-week`, `6 calendar weeks`, `all 10 of the levers`, `10 key value levers`. Either end of a range such as `6-8 weeks` or `six to eight weeks` counts. The unit before a range from 1 states the high end, so `weeks 1 to 6` is 6 weeks, and `weeks 3 to 6` is no duration. A label that ends in the count states it, so `Levers covered: all 10` and `Levers: 10` are 10 levers, but `Levers covered: all 10 of the imperatives` is not. Weeks may be written `wk` or `wks`, so `8 wks` is 8 weeks |
| `chart` | `410` | a value in the slide's chart data, as python-pptx reads it from the chart part. It takes no `where` |
| `text` | a string or a list | the load-bearing concept, with the few wordings it needs, such as `["cap", "limit"]`. A `#` stands for a count of what follows it, in digits or words up to twenty, with at most one word between, so `"# consultants"` matches `2 consultants`, `two consultants`, and `2 senior consultants`. It does not match a label's number or a date, so `Step 1 run by consultants`, `Phase 1 consultants`, `Day 1 workshop with the consultants`, `December 2 consultants`, and `2 December by the consultants` count nobody. A number after step, phase, day, week, stage, wave, workshop, month, or a month name is a label. A ` ... ` stands for up to three words and the punctuation around them, so `"team ... tbc"` matches `Team: TBC` and `Team (TBC)` |

A `require` passes when the fact is on the slide. A `forbid` or `absent` fails when it is there. Every fact reads only the text and charts of shapes that are not hidden and overlap the slide, the same shapes `lost` counts. A `where: deck` forbid also reads speaker notes. So a `require`, a slide or title `forbid`, and an `absent` fact never look at notes. Matching ignores case, Unicode composition, runs of whitespace, and curly quotes. A match cannot sit inside a longer word or number, so `39%` matches `+39%` but not `139%` or `1.39%`, and `day` does not match `days`. `evals/test_facts.py` lists the cases.

A dash or "to" between two numbers is a range, and a range states both ends. So "rose from 39 to 37%" states 39% as well. That is right for a forbid, because the old value is still on the slide. A money range needs the currency in front, so `380-410m$` is read as $410M only.

The parser does not read numbers spelled out past twenty, such as `forty-eight percent`. A fact written that way is not found, so a `require` fails and a `forbid` passes. The intent checker sees the wording either way.

Anything that is really about wording, such as tone, which phrase was used, or where a bullet sits, belongs in `intent_checks`.

## What the script decides and what it defers

The script decides these checks, in this order:

1. `source` checks that the source deck still has its pinned hash and that the output is not the source file.
2. `structure` maps every output slide to a source slide by slide id, or to an added slide in order of appearance. It reports a kept slide that is missing, a deleted slide that is still there, a new slide that no change asks for, an added slide that is missing, and a slide out of order. When two slides could explain one displacement, it blames the slide a change moved or added.
3. `scope` runs on every slide the mapping pairs, even when `structure` failed. A slide no change edits must look the same: the same text, the same slide XML, and the same related parts (charts and their embedded workbooks, images and other media, notes), followed recursively. A notes page with no text counts as no notes page. A changed slide gets `guessed` when an ambiguous non-change names it, `non-change` when a not-a-change names it, and `scope` otherwise. On an edited slide, a chart the slide keeps whose plot type changed fails `scope`, as in `c1 does not ask to change the type of slide 10's chart, but barChart became bar3DChart`. An edit target that did not change fails `missing`. A change to the deck's slide layouts, masters, or themes is one `scope` failure.
4. `lost` checks that an edited slide kept three things its edit did not target: the clauses of its on-slide text, its non-text pieces, and its chart values. Nothing else is preserved by script. Formatting, colour, size, position within the slide, stacking order, alt text, speaker notes, chart categories, and the chart's workbook can all change on an edited slide without failing `lost`, though a change that breaks the house style still fails `style`.
   - Text is compared as clauses, not lines. Each paragraph is split at `.`, `;`, `:`, `!`, `?`, and `…` before a space, and a list marker such as `1.` or `a)` is dropped. A line break only wraps a clause, so it ends none. A dot after a word of up to three characters or one with a dot inside it ends nothing, so `ca. 50%` and `U.S. Census` stay whole. A source clause survives when one output shape still says its words in order, with at most three new words between any two of them. Case, punctuation at the ends of words, paragraph and line breaks, bullet levels, and which shape the text sits in do not count. So a title split into two paragraphs, footnotes renumbered, a clause moved to a new box, and a clause that gains words all keep their text. A source clause that states one of the slide's plain `forbid` values is the edit's to rewrite and is exempt.
   - Only shapes a reader can find on the slide count, on both sides. A shape that the selection pane hides, or that lies wholly off the slide, keeps nothing, so footnote text parked in a box past the slide's edge or in a hidden box, or a footnote box dragged below the slide, is lost. Source text already hidden or off the slide is not required.
   - Pictures, charts, tables, embedded objects, and groups must survive too. Each is matched by its kind and a signature the edit does not change. A picture is matched by its image's hash, a chart by its number of series, a table by its rows and columns, an embedded object by its type, and a group by the kinds of its members. So a chart rebuilt with a named series, as `add_series("Market size ($m)", ...)` writes it, is the same chart. One that is deleted, hidden, or moved wholly off the slide, is lost, as in `c1 does not ask to remove these from slide 5, but they are gone: object 'Object 12'`. Shapes with no text that are none of these, such as lines and freeforms, are not tracked.
   - A chart's values must survive too. Every value the slide's source charts draw must still be drawn on the output slide, as many times as before, unless a `chart` forbid on the slide names it. Only charts a reader can find on the slide count, on both sides, as for text, so a correct copy of the chart parked off the slide keeps nothing. So a chart rebuilt with `replace_data` under new category names keeps its values, but a bar overwritten with the new figure or dropped from the chart is lost, as in `c1 does not ask to change these chart values on slide 10, but they are gone: 2000`. Categories, series names, series order, and which chart draws a value are not compared. A source chart that python-pptx cannot read gives no values, so its values go unchecked by script, and no failure says so.
   - Nothing an edited slide gains fails here. New text, new shapes, and new footnotes are intent scope (see below).
5. `missing` and `forbidden` check the `require` and `forbid` facts on each changed slide, every `where: deck` forbid on every slide and its speaker notes, as in `c1: 'tripled' is in the speaker notes of slide 15; it was abandoned after 00:06:54`, and the `absent` facts across the deck. A `chart` fact on a slide whose chart python-pptx cannot read (3D bar, line, or pie, stock, surface, or bar-of-pie) gets `unreadable` instead, as in `[unreadable] slide 1 chart: bar3DChart cannot be read; check by hand`. A chart that no `chart` fact reads does not stop scoring. Scope compares its part like any other. The loader rejects a scenario whose `chart` fact targets such a chart on the source slide. A maker who redraws an edited slide's chart in another plot type fails `scope`, and that slide's chart facts are not reported as `unreadable`. So `unreadable` is left for an `absent` chart fact on a slide whose source chart cannot be read. In `score.json`, a deck-wide forbid's fact has `where: notes` when its first hit is in speaker notes, and `where: deck` otherwise.
6. `layout` checks that each added slide uses its declared layout.
7. `style` reports house-style violations that the source deck did not already have, matched through the slide mapping.

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

The intent checker judges the rest. That covers every `intent_checks` line, such as tone, wording, and placement, and whether each question in `flags.json` is the right one. `score.json` lists them under `deferred`. For an ambiguous ask, the script decides only that the maker did not guess, and reports whether `flags.json` raised it.

### Additions are intent scope

The script never judges what an edited or added slide gains. A new sentence, a new text box, a new footnote, or a bullet split in two all pass `scope` and `lost`. They still meet `forbidden`, `absent`, and `style`, so a gained line that brings back an old value or breaks the house style fails. Whether the addition belongs is for the intent checker, and each scenario's `intent_checks` say what each edited or added slide may gain and what it must not.

An earlier scorer tried to decide this line by line and failed in both directions. It failed correct decks, such as a title split into two paragraphs or a survey footnote in a new box, and it passed wrong ones, such as footnote 2 replaced. These checks moved from the script to intent checks. Each row names a variant that the script now passes, which `prove.py` ties to its intent check:

| former script check | variants the script now passes | intent check that covers it |
|---|---|---|
| insurance slide 5's Footnote could gain only text that says "survey" (`may_change`) | `lunch_added_to_footnote`, `lunch_added_in_a_text_box` | c1: "Slide 5 adds nothing unrelated to the survey citation, ..., and an extended footnote gains only words about the survey." |
| a new line on an edited slide failed `scope` | rcc `split_into_two_bullets`, `split_into_bullet_and_sub_bullet` | c1: "The reworded bullet replaces "Flexibility to teams" in the same place ..., not as an extra bullet and not split into a bullet and a sub-bullet." |
| fmcg n4 read team placeholders and words (`team ... tbc`, `staffing ... tbc`, `resourcing ... to be confirmed`, `team size`, `headcount`, `FTE`) | `team_tbc_with_lars`, `staffing_tbc_with_lars`, `resourcing_to_be_confirmed`, `placeholder_for_team`, and `number_of_consultants_tbc`, `people_tbc`, `staffing_to_be_agreed_with_lars`, `team_of_x_consultants`, which no script check caught before either | n4: "The new slide holds no team-size placeholder or mention, such as ..." |

`tbd_team_bullet` ("Team size TBD") still fails by script, on the house style's placeholder-text rule, and the n4 check covers its team-size mention. n4 keeps in the script only a count of people that a script reads without doubt. That is a number right before consultants, people, persons, staff, or FTE, or with one word between, and `team of N`.

Four decks were intent scope before this change too, and now have variants. Insurance `footnote_extended_with_unrelated_words` extends footnote 1 with the survey and a remark about the venue, which the old `adds` rule let through because the extension said "survey". Solar `contents_hedged_up_to_11b`, `header_grows_to_3bn`, and `title_period_rolled_forward` each rewrite the clause that holds the old figure, which is the edit's to rewrite. The insurance additions check and the solar checks on slide 2's sentence, the header's $2bn, and the title's 2022-27 period judge them.

Text that is on the slide but cannot be seen still counts as kept. White text, text in a colour close to the background, and text behind another shape are intent scope. The script cannot decide this reliably with python-pptx. In the corpus decks, 13% of text runs inherit their colour and another 23% name a theme colour, 511 of 521 slides inherit their background, and whether one shape covers another depends on fills and stacking order that only a render shows. So the insurance scenario's c1 check says no text on slide 5 is hidden, and `footnote_2_kept_in_white_behind_curve` is its variant.

A `chart` fact reads the chart's cached values, the numbers PowerPoint draws. The chart's embedded workbook is what PowerPoint reloads on Edit Data. A maker that updates the cache but leaves the workbook stale shows the right bar until someone opens Edit Data, which brings the old number back. The script does not read workbooks, so whether the workbook matches the cache is intent-checker scope, and a scenario that changes chart data says so in its `intent_checks`.

## Private scenarios

A scenario built on a client deck lives in `$GOLDEN_PRIVATE_DIR/<name>/` and never enters the repo. It names its deck by file and hash:

```yaml
deck: {file: input.pptx, sha256: <sha256 of input.pptx>}
```

The loader rejects a `file:` deck in a scenario inside the repo, except under the repo's `private/` folder, which git ignores. Only `score.py` looks in `$GOLDEN_PRIVATE_DIR`. A name found in both `evals/` and `$GOLDEN_PRIVATE_DIR` is an error.
