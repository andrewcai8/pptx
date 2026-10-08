# Audit slides

A reviewer agent looks at each rendered slide PNG, judges it against the eight checks in `audit/rubric.md`, and writes `audit.json`. `audit_check.py` then proves the file covers every slide and follows the rubric's rules. `check` counts rules. The audit catches what a person sees, such as a title that names a topic instead of stating the takeaway, labels that overlap, or a footnote that went missing. The audit is advisory. It reports problems and never blocks a deck.

## Sub-features

- `audit-fonts` makes `deckcheck render` write `fonts.json`. It maps each font the deck uses to the family that `fc-match` finds, with `substituted` true or false. A metric-compatible clone, such as Carlito for Calibri, counts as the real font. `render` prints one `substituted: <font> -> <family>` line per substituted font, and exits 3 with an install hint when `fc-match` is missing.
- `audit-rubric` defines checks C1 to C8. C1 action-title, C2 one-message, C3 no-overlap, C4 no-clipping, C5 aligned-grid, C6 chart-legible, C7 source-on-data, and C8 title-consistent. C1 stays strict. A title must state the takeaway, and only biography, agenda, contents, and section divider slides are `n/a`.
- `audit-advisory` makes C3 and C4 advisory on every slide of a deck when `fonts.json` shows any substituted font, because a stand-in font with different widths moves wraps and edges. An advisory fail does not make a slide `needs-work`.
- `audit-schema` is one `audit.json` entry per slide PNG, `{slide, checks: {C1..C8: pass|fail|n/a}, evidence: {check: text}, advisory: [check ids], verdict: good|needs-work}`.
- `audit-check` prints `AUDIT VALID (N slides, K needs-work, A advisory)` and exits 0. A is the number of slides with a failing advisory check. Otherwise it prints `AUDIT INVALID: <reason>` and exits 1. It rejects a slide with no entry or two entries, a missing check, an illegal value, a fail with no evidence, an `advisory` list that does not match `fonts.json`, and a verdict that disagrees with the checks.
- `audit-reliability` was measured on 64 labelled BCG slides with 4 blind reviewers on an earlier copy of this rubric. Reviewers caught planted defects 55 times out of 56, and reviewer agreement was κ 0.83 to 0.86. The one miss was a 10pt grid offset. C5 does not see offsets under about 10pt. 19 of the 35 unmodified slides drew a flag. Ten were C1 flags on topic-label titles, which are fair under BCG's own convention. Six were C3 or C4 flags caused by font substitution, and three were C7 flags. With the real fonts installed, 5 of the 6 font flags cleared.

## How to get to it (user POV)

- Prove a changed deck with `SKILL.md` Drive, then audit the changed slides after `render`.
- Read the needs-work slides in `audit.json`, open each PNG, and decide whether to fix the slide. A fail is a reason to look, not a gate.

## Driving it with deckcheck

Preconditions:

- Baseline preconditions from `README.md` hold, and `doctor` finds `soffice`.
- The deck's fonts are installed. On Debian or Ubuntu, accept the Microsoft fonts EULA and install Trebuchet MS and the Calibri stand-in:

  ```bash
  echo ttf-mscorefonts-installer msttcorefonts/accepted-mscorefonts-eula select true | sudo debconf-set-selections
  sudo apt-get install -y ttf-mscorefonts-installer fonts-crosextra-carlito
  ```

- `fc-match` exists. It ships with fontconfig, and `brew install fontconfig` adds it on macOS.

Steps:

- **Render.** Run `uv run --project deckcheck deckcheck render $RUN/new.pptx --out $RUN/render`. Exit 0. Read any `substituted:` lines, and open `$RUN/render/fonts.json`. A substituted font makes C3 and C4 advisory for the whole deck.
- **Pick the slides.** Audit the slides that `diff` lists as `changed` or `added`. Copy their PNGs and `fonts.json` into one directory, for example `mkdir -p $RUN/audit/slides && cp $RUN/render/fonts.json $RUN/render/slide-2.png $RUN/render/slide-5.png $RUN/audit/slides/`. To audit every slide, use `$RUN/render` as the slide directory instead.
- **Review blind.** Spawn one fresh reviewer subagent per deck. Give it only the three paths below, never the deck, the diff, the meeting, or the pipeline's prompt. A reviewer that knows what the slide was meant to say reads that intent into the image.

  ```text
  Read <repo>/.claude/skills/verify-pptx/audit/rubric.md. Then read fonts.json and every slide-N.png in <slide-dir>, and nothing else. Audit each PNG against the rubric and write the JSON array it describes to <run>/audit/audit.json.
  ```

- **Validate.** Run `uv run --project deckcheck python .claude/skills/verify-pptx/scripts/audit_check.py $RUN/audit/slides $RUN/audit/audit.json`. Require `AUDIT VALID`. On `AUDIT INVALID`, give the reason to a fresh reviewer with the same prompt and validate again.
- **Report.** List each needs-work slide with its failed checks and evidence. Open its PNG and say whether each flag is fair. List advisory fails apart from the rest, as layout to confirm in PowerPoint.

## Gotchas

- Henderson BCG Sans is BCG's brand font and is not public. `fc-match` resolves it to a fallback, Verdana on a Debian machine with the Microsoft fonts. A deck set in it always gets advisory C3 and C4. In the corpus, Henderson BCG Sans appears only in font slots that carry no visible Latin text, so no corpus deck lists it in `fonts.json`.
- LibreOffice may not honour shrink-to-fit. Text that PowerPoint shrinks to fit its box can overflow in the render and draw a C4 fail. Check the slide in PowerPoint before you act on it.
- `fonts.json` lists the fonts that `check` resolves, the explicit run fonts and the theme fonts. Fonts set only in master or layout text styles, charts, or bullets are not listed, so their substitution goes unreported.
- Advisory applies to the whole deck. One substituted font makes C3 and C4 advisory on every slide, including slides that never use it.
- A substituted font can also add a line to a heading and push one column of a grid down, which fails C5. C5 is never advisory, so when `fonts.json` shows a substitution, check a C5 fail against the wraps before you act on it.
- Renders are 80 dpi. C5 misses offsets under about 10pt, and C6 judges legibility at that resolution, not at full screen.
- The reliability numbers hold for the rubric as measured. Measure a rubric change against labelled slides again before you trust it.
