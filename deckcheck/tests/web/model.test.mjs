import assert from "node:assert/strict";
import { test } from "node:test";

import { chipLabel, choices, deck, editable, marks, segments, slideLabel, viewOf } from "../../src/deckcheck/review/web/model.js";

const SLIDE = { w: 12192000, h: 6858000 };
const TITLE = { id: 2, name: "Title 2", box: { x: 630000, y: 622800, w: 10933350, h: 664797 } };

const change = (id, kind, slide, extra = {}) => ({
  id,
  kind,
  slide,
  structural: !["replace_text", "insert_paragraph"].includes(kind),
  shape: null,
  before: "",
  after: "",
  span: null,
  depends_on: null,
  admits: ["keep_new", "keep_old"],
  ...extra,
});

const review = {
  flags: [{ id: "f1", question: "Which year?", slides: [13], refs: [] }],
  held: [{ id: "h1", text: "Keep the wording.", slides: [14], refs: [] }],
  slides: [
    { key: 10, source_index: 1, executed_index: 1, title: "Intro" },
    { key: 11, source_index: 2, executed_index: null, title: "Credentials" },
    { key: 12, source_index: 3, executed_index: 2, title: "Overview" },
    { key: 13, source_index: 4, executed_index: 4, title: "Market" },
    { key: 14, source_index: 5, executed_index: 5, title: "Generations" },
    { key: 15, source_index: 6, executed_index: 3, title: "Innovation" },
    { key: "add", source_index: null, executed_index: 6, title: "Next steps" },
  ],
  changes: [
    change("drop", "delete_slide", 11),
    change("move", "move_slide", 15),
    change("m1", "replace_text", 13, { shape: TITLE }),
    change("m2", "replace_text", 13, { shape: TITLE }),
    change("cell", "set_cell", 13, { shape: { id: 3, name: "Table 39", box: { x: 0, y: 3429000, w: 6096000, h: 1714500 } } }),
    change("t1", "replace_text", 14, { shape: { id: 5, name: "Title 1", box: TITLE.box } }),
    change("t2", "insert_paragraph", 14, { shape: { id: 6, name: "Footnote", box: TITLE.box } }),
    change("add", "add_slide", "add"),
    change("fill", "fill_placeholder", "add", { shape: { id: 2, name: "Title 1", box: null }, depends_on: "add" }),
  ],
};

test("the rail follows the new deck and puts a deleted slide where it was", () => {
  const rail = deck(review, { t1: "keep_new" });

  assert.deepEqual(
    rail.map((e) => [e.token, e.view, e.frame, e.changes.length, e.pending, e.oldIndex, e.newIndex]),
    [
      ["s10", "plain", null, 0, 0, 1, 1],
      ["s11", "structural", "deleted", 1, 1, 2, null],
      ["s12", "plain", null, 0, 0, 3, 2],
      ["s15", "structural", "moved", 1, 1, 6, 3],
      ["s13", "structural", null, 3, 3, 4, 4],
      ["s14", "text", null, 2, 1, 5, 5],
      ["a-add", "structural", "added", 2, 2, null, 6],
    ],
  );
  assert.deepEqual(rail[4].flags.map((f) => f.id), ["f1"]);
  assert.deepEqual(rail[5].held.map((h) => h.id), ["h1"]);
});

test("a deleted first slide leads the rail", () => {
  const slides = [
    { key: 1, source_index: 1, executed_index: null, title: "Cover" },
    { key: 2, source_index: 2, executed_index: null, title: "Agenda" },
    { key: 3, source_index: 3, executed_index: 1, title: "Body" },
  ];

  assert.deepEqual(deck({ ...review, slides, changes: [] }, {}).map((e) => e.token), ["s1", "s2", "s3"]);
});

test("a changed slide is labelled by its place, not by the source title", () => {
  const slides = [
    { key: 1, source_index: 1, executed_index: null, title: "Cover" },
    { key: 2, source_index: 2, executed_index: 1, title: "About 50% of clients" },
    { key: 3, source_index: 3, executed_index: 2, title: "Body" },
  ];
  const shifted = { ...review, slides, changes: [change("drop", "delete_slide", 1), change("r", "replace_text", 2)] };

  assert.deepEqual(
    [...deck(shifted, {}), ...deck(review, {})].map((e) => [slideLabel(e), e.title]),
    [
      ["Slide 1 of the old deck, deleted", null],
      ["Slide 1, slide 2 in the old deck", null],
      ["Slide 2, slide 3 in the old deck", "Body"],
      ["Slide 1", "Intro"],
      ["Slide 2 of the old deck, deleted", null],
      ["Slide 2, slide 3 in the old deck", "Overview"],
      ["Slide 3, moved from slide 6 to slide 3", null],
      ["Slide 4", null],
      ["Slide 5", null],
      ["New slide 6", null],
    ],
  );
});

test("a slide chip names the old number the meeting used when it differs", () => {
  assert.deepEqual(deck(review, {}).map(chipLabel), [
    "Slide 1",
    "Old slide 2 (deleted)",
    "Slide 2 (was 3)",
    "Slide 3 (was 6)",
    "Slide 4",
    "Slide 5",
    "New slide 6",
  ]);
});

test("two changes on one shape share one mark, placed in percent of the slide", () => {
  const onSlide = review.changes.filter((c) => c.slide === 13 || c.slide === "add");

  assert.deepEqual(
    marks(onSlide, SLIDE).map((m) => [m.shape, m.number, m.changes.map((c) => c.id), m.left, m.top, m.width, m.height]),
    [
      [2, 1, ["m1", "m2"], 5.167, 9.081, 89.676, 9.694],
      [3, 2, ["cell"], 0, 50, 50, 25],
    ],
  );
});

test("editable slices the replacement by code points", () => {
  const c = change("t", "replace_text", 1, {
    before: "📈 Revenue grew 12% in 2025",
    after: "📈 Revenue grew 15.5% in 2025",
    span: [15, 18],
    admits: ["keep_new", "keep_old", "edited"],
  });

  assert.equal(editable(c, "pending"), "15.5%");
  assert.equal(editable(c, { edited: "16%" }), "16%");
  assert.equal(editable(change("p", "fill_placeholder", "add", { after: "Raise prices\nHold discounts" }), "keep_old"), "Raise prices\nHold discounts");
});

test("a slide with a change the engine could not place is structural", () => {
  const boxed = change("t", "replace_text", 1, { shape: TITLE });
  const loose = change("u", "replace_text", 1, { shape: { id: 7, name: "Subtitle 2", box: null } });

  assert.equal(viewOf([boxed]), "text");
  assert.equal(viewOf([boxed, loose]), "structural");
  assert.equal(viewOf([change("v", "insert_paragraph", 1)]), "structural");
});

test("segments show the decision's outcome against the source text", () => {
  const long = "x".repeat(70);
  const c = change("t", "replace_text", 1, {
    before: `${long} grew 12% in 2025\vsince launch`,
    after: `${long} grew 15% in 2025\vsince launch`,
    span: [76, 79],
  });
  const tail = [{ kind: "same", text: " in 2025\nsince launch" }];
  const head = { kind: "same", text: `…${"x".repeat(54)} grew ` };

  assert.deepEqual(segments(c, "keep_new"), [head, { kind: "del", text: "12%" }, { kind: "ins", text: "15%" }, ...tail]);
  assert.deepEqual(segments(c, { edited: "16%" }), [head, { kind: "del", text: "12%" }, { kind: "ins", text: "16%" }, ...tail]);
  assert.deepEqual(segments(c, "keep_old"), [{ kind: "same", text: `…${"x".repeat(54)} grew 12% in 2025\nsince launch` }]);
  assert.deepEqual(segments(change("c", "set_cell", 1, { before: "$120m", after: "$130m" }), "pending"), [
    { kind: "del", text: "$120m" },
    { kind: "ins", text: "$130m" },
  ]);
  assert.deepEqual(segments(change("p", "insert_paragraph", 1, { after: "Costs fall" }), "keep_new"), [{ kind: "ins", text: "Costs fall" }]);
});

test("each kind of change names its own choices", () => {
  assert.deepEqual(choices(change("d", "delete_slide", 1)).map((c) => c.label), ["Keep deleted", "Restore slide"]);
  assert.deepEqual(choices(change("m", "move_slide", 1)).map((c) => c.label), ["Keep new order", "Move back"]);
  assert.deepEqual(choices(change("a", "add_slide", "a")).map((c) => c.label), ["Keep slide", "Drop slide"]);
  assert.deepEqual(
    choices(change("t", "replace_text", 1, { admits: ["keep_new", "keep_old", "edited"] })).map((c) => [c.decide, c.label]),
    [
      ["keep_new", "Keep new"],
      ["keep_old", "Keep old"],
      ["edit", "Edit myself"],
    ],
  );
});
