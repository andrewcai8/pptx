import { h } from "./dom.js";
import * as M from "./model.js";
import { call, img, meetingUrl, mount, render, state, ui } from "./store.js";

const DECIDED = { pending: "Needs a decision", keep_new: "Kept new", keep_old: "Kept old", edited: "Edited" };

export function refs(list) {
  return h(
    "ul",
    { class: "refs" },
    list.map((r) => h("li", null, h("span", { class: "t" }, r.t), " ", h("span", { class: "who" }, r.speaker), " ", h("q", null, r.quote))),
  );
}

function frame(side, src, marks, extra = {}) {
  const { w, h: height } = state.view.slide;
  return h(
    "div",
    {
      class: `frame ${extra.class ?? ""}`,
      "data-side": side,
      style: { aspectRatio: `${w} / ${height}`, maxWidth: extra.fit ? `calc((100vh - 230px) * ${w / height})` : null },
    },
    src && img(src, side),
    marks.map((m) => markEl(m, extra.onMark)),
  );
}

function markEl(m, handlers = {}) {
  const decision = M.markState(m.changes, state.decisions);
  return h(
    "div",
    {
      class: `mark ${decision}`,
      "data-mark": m.shape,
      "data-changes": m.changes.map((c) => c.id).join(" "),
      style: { left: `${m.left}%`, top: `${m.top}%`, width: `${m.width}%`, height: `${m.height}%` },
      onmouseenter: () => handlers.enter?.(m),
      onmouseleave: () => handlers.leave?.(m),
      onclick: (event) => handlers.click?.(m, event),
    },
    decision === "pending" && h("span", { class: "badge" }, m.number),
    decision === "edited" && h("span", { class: "badge" }, "✎"),
  );
}

export function plainView(e) {
  return h(
    "div",
    { class: "text-view" },
    h("div", { class: "slide-col" }, frame("new", render("new", e.newIndex), [], { fit: true })),
    h("p", { class: "muted lane" }, "Nothing changed on this slide."),
  );
}

let hideTimer = null;

export function textView(v, e) {
  const marks = M.marks(e.changes, v.slide);
  const onMark = {
    enter: (m) => {
      clearTimeout(hideTimer);
      if (state.pinned === null) showBubble(m.shape);
    },
    leave: () => {
      if (state.pinned === null) hideTimer = setTimeout(hideBubble, 250);
    },
    click: (m, event) => {
      event.stopPropagation();
      state.pinned = m.shape;
      showBubble(m.shape);
    },
  };
  return h(
    "div",
    { class: "text-view" },
    h("div", { class: "slide-col" }, frame("new", render("new", e.newIndex), marks, { onMark, fit: true })),
    h(
      "div",
      { class: "lane" },
      h("p", { class: "hint" }, "Point at a highlight to see what changed. Click it to keep the note open."),
      h("div", {
        id: "bubble",
        hidden: true,
        onclick: () => (state.pinned ??= state.open),
        onmouseenter: () => clearTimeout(hideTimer),
        onmouseleave: () => {
          if (state.pinned === null) hideTimer = setTimeout(hideBubble, 250);
        },
      }),
    ),
  );
}

export function showBubble(shape) {
  state.open = shape;
  const bubble = document.getElementById("bubble");
  const mark = document.querySelector(`#stage [data-side="new"] [data-mark="${shape}"]`);
  if (!bubble || !mark) return;
  const e = M.deck(state.view.review, state.decisions).find((x) => x.token === state.slide);
  const changes = e.changes.filter((c) => c.shape?.id === shape);
  mount(
    bubble,
    h("span", { class: "pointer" }),
    h(
      "div",
      { class: "bubble-body" },
      h(
        "div",
        { class: "bubble-head" },
        h("strong", null, changes[0].shape.name),
        state.pinned !== null && h("button", { class: "close quiet", onclick: (event) => (event.stopPropagation(), unpin()), title: "Close (Esc)" }, "×"),
      ),
      changes.map((c) => card(c, { named: false })),
    ),
  );
  bubble.hidden = false;
  bubble.classList.toggle("pinned", state.pinned !== null);
  for (const m of document.querySelectorAll("#stage .mark")) m.classList.toggle("lit", m.dataset.mark === String(shape));
  placeBubble();
}

export function placeBubble() {
  const bubble = document.getElementById("bubble");
  const mark = document.querySelector(`#stage [data-side="new"] [data-mark="${state.open}"]`);
  if (!bubble || bubble.hidden || !mark) return;
  const lane = bubble.parentElement.getBoundingClientRect();
  const r = mark.getBoundingClientRect();
  const stage = document.getElementById("stage").getBoundingClientRect();
  const centre = r.top + r.height / 2;
  const height = bubble.offsetHeight;
  const top = Math.max(stage.top + 8, Math.min(centre - height / 2, stage.bottom - height - 8));
  Object.assign(bubble.style, { left: `${lane.left}px`, width: `${lane.width}px`, top: `${top}px` });
  bubble.querySelector(".pointer").style.top = `${Math.max(14, Math.min(centre - top, height - 14))}px`;
}

export function hideBubble() {
  const bubble = document.getElementById("bubble");
  if (!bubble || state.pinned !== null) return;
  bubble.hidden = true;
  for (const m of document.querySelectorAll("#stage .mark.lit")) m.classList.remove("lit");
}

export function unpin() {
  state.pinned = null;
  state.editing = null;
  hideBubble();
}

export function structuralView(v, e) {
  const marks = M.marks(e.changes, v.slide);
  const light = (ids, on) => {
    for (const el of document.querySelectorAll("#stage [data-mark], #changes [data-change]")) {
      const mine = (el.dataset.changes ?? el.dataset.change).split(" ");
      if (ids.some((id) => mine.includes(id))) el.classList.toggle("lit", on);
    }
  };
  const onMark = {
    enter: (m) => light(m.changes.map((c) => c.id), true),
    leave: (m) => light(m.changes.map((c) => c.id), false),
    click: (m) => document.querySelector(`#changes [data-change="${CSS.escape(m.changes[0].id)}"]`)?.scrollIntoView({ block: "nearest", behavior: "smooth" }),
  };
  const old = e.frame === "added" ? frame("old", null, [], { class: "empty" }) : frame("old", render("old", e.oldIndex), marks, { onMark });
  const now =
    e.frame === "deleted" ? frame("new", null, [], { class: "empty" }) : frame("new", render("new", e.newIndex), marks, { onMark, class: e.frame === "added" ? "added" : "" });
  if (e.frame === "added") old.append(h("span", { class: "empty-label" }, "New slide"));
  if (e.frame === "deleted") now.append(h("span", { class: "empty-label" }, "Deleted"));
  return h(
    "div",
    { class: "structural-view" },
    h(
      "div",
      { class: "two-up" },
      h("div", { class: "side" }, h("div", { class: "side-label" }, "Before"), old),
      h("div", { class: "side" }, h("div", { class: "side-label" }, "After"), now),
    ),
    h(
      "aside",
      { id: "changes" },
      e.changes.map((c) => {
        const el = card(c, { number: marks.find((m) => m.changes.includes(c))?.number });
        el.addEventListener("mouseenter", () => light([c.id], true));
        el.addEventListener("mouseleave", () => light([c.id], false));
        return el;
      }),
    ),
  );
}

function card(c, { number, named = true } = {}) {
  const decision = M.decisionOf(state.decisions, c);
  const kind = M.kindOf(decision);
  const dropped = M.dropped(c, state.view.review, state.decisions);
  return h(
    "div",
    { class: `change ${kind}`, "data-change": c.id, "data-decision": kind },
    h(
      "div",
      { class: "change-head" },
      number && h("span", { class: "number" }, number),
      h("span", { class: "kind" }, M.KIND[c.kind] ?? c.kind),
      named && c.shape && h("span", { class: "muted" }, c.shape.name),
      h("span", { class: `status ${kind}` }, dropped ? "Dropped with the slide" : DECIDED[kind]),
    ),
    state.editing === c.id ? editor(c) : diff(c, decision),
    c.notes.length > 0 && h("ul", { class: "notes" }, c.notes.map((n) => h("li", null, n))),
    h("p", { class: "why" }, c.rationale),
    refs(c.refs),
    state.editing !== c.id &&
      h(
        "div",
        { class: "choices" },
        M.choices(c).map((choice) => {
          const on = choice.decide === (kind === "edited" ? "edit" : kind);
          return h(
            "button",
            {
              "data-decide": choice.decide,
              class: on ? "on" : "",
              "aria-pressed": String(on),
              disabled: dropped,
              onclick: () => (choice.decide === "edit" ? startEdit(c, decision) : decide(c.id, choice.decide)),
            },
            choice.label,
          );
        }),
      ),
  );
}

function diff(c, decision) {
  if (["add_slide", "delete_slide", "move_slide"].includes(c.kind)) {
    const words = { add_slide: [null, c.after], delete_slide: [c.before, null], move_slide: [c.before, c.after] }[c.kind];
    return h(
      "dl",
      { class: "was-now" },
      words[0] && [h("dt", null, "Before"), h("dd", null, words[0])],
      words[1] && [h("dt", null, "Now"), h("dd", null, words[1])],
    );
  }
  return [
    h("p", { class: "diff" }, M.segments(c, decision).map((s) => h(s.kind === "same" ? "span" : s.kind, null, s.text))),
    M.kindOf(decision) === "keep_old" && h("p", { class: "proposed" }, "Proposed: ", h("del", null, M.editable(c, "keep_new"))),
  ];
}

function startEdit(c, decision) {
  state.editing = c.id;
  state.drafts[c.id] ??= M.editable(c, decision);
  if (document.getElementById("stage").dataset.view === "text") state.pinned = c.shape.id;
  ui.rerender();
  document.querySelector(`textarea[data-edit="${CSS.escape(c.id)}"]`)?.focus();
}

function editor(c) {
  const around = M.context(c);
  const draft = state.drafts[c.id] ?? "";
  return h(
    "div",
    { class: "editor", onclick: (event) => event.stopPropagation() },
    around.before && h("span", { class: "context" }, around.before),
    h("textarea", {
      "data-edit": c.id,
      rows: Math.min(8, Math.max(2, draft.split("\n").length + 1)),
      value: draft,
      oninput: (event) => (state.drafts[c.id] = event.target.value),
      onkeydown: (event) => {
        if (event.key === "Enter" && (event.metaKey || event.ctrlKey)) save(c);
      },
    }),
    around.after && h("span", { class: "context" }, around.after),
    state.errors[c.id] && h("p", { class: "error", "data-error": "" }, state.errors[c.id]),
    h(
      "div",
      { class: "choices" },
      h("button", { class: "primary", "data-act": "save", onclick: () => save(c) }, "Save"),
      h("button", { "data-act": "cancel", onclick: () => cancelEdit(c) }, "Cancel"),
    ),
  );
}

async function save(c) {
  const text = document.querySelector(`textarea[data-edit="${CSS.escape(c.id)}"]`)?.value ?? state.drafts[c.id];
  state.drafts[c.id] = text;
  if (await decide(c.id, M.edited(c, text))) {
    state.editing = null;
    delete state.drafts[c.id];
    delete state.errors[c.id];
    ui.rerender();
  }
}

export function cancelEdit(c) {
  state.editing = null;
  delete state.drafts[c.id];
  delete state.errors[c.id];
  ui.rerender();
}

async function decide(cid, decision) {
  const { status, body } = await call("POST", `${meetingUrl(state.view.id)}/decisions`, { decisions: { [cid]: decision } });
  if (status === 200) {
    if (state.final && !body.final) state.stale = true;
    state.decisions = body.decisions;
    state.final = body.final;
    delete state.errors[cid];
    state.notice = null;
  } else if (status === 422) {
    state.errors[cid] = body.problems.map((p) => p.message).join(" ");
  } else {
    state.notice = body.error ?? `The server answered ${status}.`;
  }
  ui.rerender();
  return status === 200;
}
