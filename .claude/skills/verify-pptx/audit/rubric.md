# Slide visual audit rubric

Audit each slide PNG you were given. Judge only what the PNG shows. Fail only on something you can point to in the image, and name it in the evidence. When unsure, pass.

Read `fonts.json` next to the PNGs before you judge. If any font in it has `"substituted": true`, the render used a stand-in font whose different widths move wraps and edges. C3, C4, and C5 are then advisory on every slide of the deck. Still judge them and report what you see.

Each check returns `pass`, `fail`, or `n/a`. Only C1, C5, C6, and C7 may be `n/a`, and only where the check says so. C2, C3, C4, and C8 are always `pass` or `fail`.

**C1 action-title.** Pass when the title asserts a finding or recommendation about its subject: a sentence or clause with a main verb, or a quantified or comparative assertion. A kicker before a colon ("Looking back: value creation has lagged") is ignored, and the part after it is judged. Fail when the title only names the subject, including a noun phrase carrying a relative clause ("Seven practices leading retailers are considering", "Impact scores by seniority level", "Layout changes"). `n/a` for cover, biography, agenda, contents, and section divider slides.

**C2 one-message.** Pass when every chart, table, and text block supports or details the title's claim. Several charts on the same subject pass. Fail when a block presents a subject the title does not cover and that does not feed the claim.

**C3 no-overlap.** Fail when text overlaps other text, a chart's bars, axis, or plot area, or an image, so that characters or data marks are covered or struck through. Text placed on purpose inside its own label box, callout, or bar segment is not overlap. A callout that covers other data labels is overlap.

**C4 no-clipping.** Fail when text is cut off by the slide edge or by its container (missing letters, a line ending mid-word at a box border, text spilling past a box border).

**C5 aligned-grid.** Applies when the slide has a repeated set of three or more parallel elements. Fail when a member's left, right, or top edge is out of line with the others, or its spacing differs by more than about a quarter of the gap between members. `n/a` when there is no repeated set. Offsets under about 10pt are below what this check reliably sees at 80 dpi.

**C6 chart-legible.** Applies when the slide has a chart. Fail when the chart's values cannot be read at this resolution (data labels and axis tick labels both missing or too small to read), or when no unit or measure is stated anywhere for the chart. `n/a` when there is no chart.

**C7 source-on-data.** Applies when the slide shows quantitative data (a chart, a table of figures, or numeric claims). Pass when a source or note line is present, usually small text at the bottom left. Inline "(Link)" references alone do not pass. `n/a` when there is no quantitative data.

**C8 title-consistent.** Fail when something in the visual contradicts the title: a direction (title says falling, chart rises) or a number (title says $3.8bn, chart says $380m).

Verdict: `needs-work` when any check outside `advisory` fails, otherwise `good`. An advisory fail alone leaves the slide `good`.

Write one JSON array to the `audit.json` path you were given, with one entry per PNG. `slide` is the N in `slide-N.png`. `evidence` holds one entry for every failed check. `advisory` is `["C3", "C4", "C5"]` when `fonts.json` shows a substituted font, and `[]` otherwise.

```json
[
  {
    "slide": 3,
    "checks": {"C1": "fail", "C2": "pass", "C3": "fail", "C4": "pass", "C5": "n/a", "C6": "pass", "C7": "pass", "C8": "pass"},
    "evidence": {"C1": "title 'Market overview' names the subject only", "C3": "the 2024 data label overlaps the legend"},
    "advisory": ["C3", "C4", "C5"],
    "verdict": "needs-work"
  }
]
```
