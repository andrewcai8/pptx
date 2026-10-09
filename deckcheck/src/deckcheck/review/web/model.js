const CONTEXT = 60;

export const KIND = {
  replace_text: "Text",
  insert_paragraph: "New paragraph",
  set_cell: "Table cell",
  set_chart_value: "Chart value",
  add_slide: "New slide",
  fill_placeholder: "Placeholder",
  delete_slide: "Deleted slide",
  move_slide: "Moved slide",
};

const WORDS = {
  default: ["Keep new", "Keep old"],
  add_slide: ["Keep slide", "Drop slide"],
  delete_slide: ["Keep deleted", "Restore slide"],
  move_slide: ["Keep new order", "Move back"],
};

export const token = (key) => (typeof key === "number" ? `s${key}` : `a-${key}`);

export const decisionOf = (decisions, change) => decisions[change.id] ?? "pending";
export const kindOf = (decision) => (typeof decision === "string" ? decision : "edited");

export function deck(review, decisions) {
  const byToken = new Map();
  for (const c of review.changes) {
    const t = token(c.slide);
    byToken.set(t, [...(byToken.get(t) ?? []), c]);
  }
  const order = review.slides.filter((s) => s.executed_index !== null).sort((a, b) => a.executed_index - b.executed_index);
  const ghosts = review.slides.filter((s) => s.executed_index === null).sort((a, b) => a.source_index - b.source_index);
  for (const ghost of ghosts) {
    const anchors = order.filter((s) => s.source_index !== null && s.executed_index !== null && s.source_index < ghost.source_index);
    const anchor = anchors.sort((a, b) => b.source_index - a.source_index)[0];
    let at = anchor ? order.indexOf(anchor) + 1 : 0;
    while (at < order.length && order[at].executed_index === null && order[at].source_index < ghost.source_index) at++;
    order.splice(at, 0, ghost);
  }
  return order.map((s) => {
    const t = token(s.key);
    const changes = byToken.get(t) ?? [];
    const kinds = new Set(changes.map((c) => c.kind));
    const frame =
      s.source_index === null ? "added" : s.executed_index === null ? "deleted" : kinds.has("move_slide") ? "moved" : null;
    return {
      token: t,
      key: s.key,
      title: changes.length ? null : s.title,
      view: viewOf(changes),
      frame,
      changes,
      pending: changes.filter((c) => decisionOf(decisions, c) === "pending").length,
      flags: review.flags.filter((f) => f.slides.includes(s.key)),
      held: review.held.filter((h) => h.slides.includes(s.key)),
      oldIndex: s.source_index,
      newIndex: s.executed_index,
    };
  });
}

export function slideLabel(e) {
  if (e.frame === "added") return `New slide ${e.newIndex}`;
  if (e.frame === "deleted") return `Slide ${e.oldIndex} of the old deck, deleted`;
  if (e.frame === "moved") return `Slide ${e.newIndex}, moved from slide ${e.oldIndex} to slide ${e.newIndex}`;
  return e.oldIndex === e.newIndex ? `Slide ${e.newIndex}` : `Slide ${e.newIndex}, slide ${e.oldIndex} in the old deck`;
}

export function chipLabel(e) {
  if (e.frame === "added") return `New slide ${e.newIndex}`;
  if (e.frame === "deleted") return `Old slide ${e.oldIndex} (deleted)`;
  return e.oldIndex === e.newIndex ? `Slide ${e.newIndex}` : `Slide ${e.newIndex} (was ${e.oldIndex})`;
}

export function viewOf(changes) {
  if (changes.length === 0) return "plain";
  return changes.some((c) => c.structural) ? "structural" : "text";
}

const pct = (part, whole) => Math.round((part / whole) * 100000) / 1000;

export function marks(changes, slide) {
  const byShape = new Map();
  for (const c of changes) {
    if (!c.shape?.box) continue;
    const mark = byShape.get(c.shape.id) ?? { shape: c.shape.id, name: c.shape.name, box: c.shape.box, changes: [] };
    mark.changes.push(c);
    byShape.set(c.shape.id, mark);
  }
  return [...byShape.values()].map((m, i) => ({
    shape: m.shape,
    name: m.name,
    number: i + 1,
    changes: m.changes,
    left: pct(m.box.x, slide.w),
    top: pct(m.box.y, slide.h),
    width: pct(m.box.w, slide.w),
    height: pct(m.box.h, slide.h),
  }));
}

export function markState(changes, decisions) {
  const kinds = changes.map((c) => kindOf(decisionOf(decisions, c)));
  if (kinds.includes("pending")) return "pending";
  if (kinds.every((k) => k === "keep_old")) return "keep_old";
  if (kinds.includes("edited")) return "edited";
  return "keep_new";
}

const points = (s) => [...s];

function replacement(change) {
  const [s, e] = change.span;
  const before = points(change.before);
  const after = points(change.after);
  return after.slice(s, after.length - (before.length - e)).join("");
}

export function editable(change, decision) {
  if (kindOf(decision) === "edited") return decision.edited;
  return change.kind === "replace_text" ? replacement(change) : change.after ?? "";
}

function outcome(change, decision) {
  const kind = kindOf(decision);
  if (kind === "edited") return decision.edited;
  if (kind === "keep_old") return change.kind === "replace_text" ? quote(change) : change.before ?? "";
  return change.kind === "replace_text" ? replacement(change) : change.after ?? "";
}

function quote(change) {
  const [s, e] = change.span;
  return points(change.before).slice(s, e).join("");
}

const shown = (text) => text.replaceAll("\v", "\n");

export function context(change) {
  if (change.kind !== "replace_text") return { before: "", after: "" };
  const [s, e] = change.span;
  const text = points(change.before);
  const head = text.slice(0, s);
  const tail = text.slice(e);
  return {
    before: shown((head.length > CONTEXT ? "…" : "") + head.slice(-CONTEXT).join("")),
    after: shown(tail.slice(0, CONTEXT).join("") + (tail.length > CONTEXT ? "…" : "")),
  };
}

export function segments(change, decision) {
  const now = outcome(change, decision);
  const old = change.kind === "replace_text" ? quote(change) : change.before ?? "";
  const { before, after } = context(change);
  const middle =
    old === now
      ? [{ kind: "same", text: shown(old) }]
      : [
          { kind: "del", text: shown(old) },
          { kind: "ins", text: shown(now) },
        ];
  const out = [];
  for (const seg of [{ kind: "same", text: before }, ...middle, { kind: "same", text: after }]) {
    if (seg.text === "") continue;
    const last = out.at(-1);
    if (last?.kind === "same" && seg.kind === "same") last.text += seg.text;
    else out.push({ ...seg });
  }
  return out;
}

export function choices(change) {
  const [keep, revert] = WORDS[change.kind] ?? WORDS.default;
  const out = [
    { decide: "keep_new", label: keep },
    { decide: "keep_old", label: revert },
  ];
  if (change.admits.includes("edited")) out.push({ decide: "edit", label: "Edit myself" });
  return out;
}

export function dropped(change, review, decisions) {
  if (change.kind !== "fill_placeholder" || change.depends_on === null) return false;
  const add = review.changes.find((c) => c.id === change.depends_on);
  return add !== undefined && decisionOf(decisions, add) === "keep_old";
}

export function tally(review, decisions) {
  const total = review.changes.length;
  const pending = review.changes.filter((c) => decisionOf(decisions, c) === "pending").length;
  return { total, decided: total - pending, pending };
}

export const ref = (r) => `${r.t} ${r.speaker} “${r.quote}”`;
