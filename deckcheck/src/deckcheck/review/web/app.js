import { h } from "./dom.js";
import * as M from "./model.js";
import { cancelEdit, hideBubble, placeBubble, plainView, refs, showBubble, structuralView, textView, unpin } from "./slide.js";
import { call, img, meetingUrl, mount, render, state, ui } from "./store.js";

const STEPS = { making: "Making changes", executing: "Writing the deck", rendering: "Rendering slides, about 40 s" };

const root = document.getElementById("app");

const plural = (n, one, many) => `${n} ${n === 1 ? one : many}`;

function parseHash() {
  const parts = location.hash.replace(/^#\/?/, "").split("/").filter(Boolean).map(decodeURIComponent);
  if (parts[0] === "m" && parts.length >= 3) return { mid: `${parts[1]}/${parts[2]}`, slide: parts[3] ?? null };
  return { mid: null };
}

async function route() {
  const { mid, slide } = parseHash();
  clearInterval(state.poll);
  hideBubble();
  if (mid === null) {
    state.view = null;
    await refreshMeetings();
    state.poll = setInterval(refreshMeetings, 1000);
    return;
  }
  if (state.view?.id !== mid) await openMeeting(mid);
  if (!state.view) return renderMissing();
  const rail = M.deck(state.view.review, state.decisions);
  const chosen = rail.find((e) => e.token === slide) ?? rail.find((e) => e.changes.length) ?? rail[0];
  select(chosen.token);
}

async function openMeeting(mid) {
  const { status, body } = await call("GET", meetingUrl(mid));
  state.view = status === 200 ? body : null;
  state.missing = status === 200 ? null : body;
  if (!state.view) return;
  state.decisions = Object.fromEntries(body.review.changes.map((c) => [c.id, c.decision]));
  state.final = body.final;
  Object.assign(state, { stale: false, pinned: null, editing: null, drafts: {}, errors: {}, notice: null, published: false });
}

function select(token) {
  state.slide = token;
  state.pinned = null;
  state.editing = null;
  hideBubble();
  history.replaceState(null, "", `#/m/${state.view.id}/${encodeURIComponent(token)}`);
  renderReview();
  document.querySelector(`#rail [data-slide="${CSS.escape(token)}"]`)?.scrollIntoView({ block: "nearest" });
}

async function refreshMeetings() {
  const { body } = await call("GET", "/api/meetings");
  const before = JSON.stringify(state.meetings);
  state.meetings = body.meetings;
  for (const m of state.meetings) {
    if (!state.mine.has(m.id) || m.state.is === "processing") continue;
    state.mine.delete(m.id);
    if (m.state.is === "ready" && parseHash().mid === null) {
      location.hash = `#/m/${m.id}`;
      return;
    }
  }
  if (parseHash().mid === null && (before !== JSON.stringify(state.meetings) || !root.querySelector(".home"))) renderHome();
}

async function processMeeting(m, again) {
  if (again && !confirm("Process this meeting again? Your decisions on it are discarded.")) return;
  const { status, body } = await call("POST", `${meetingUrl(m.id)}/process`, again ? { again: true } : {});
  if (status === 202) {
    state.mine.add(m.id);
    state.meetings = state.meetings.map((x) => (x.id === m.id ? body : x));
  }
  renderHome();
}

function renderHome() {
  mount(
    root,
    h("header", { class: "top" }, h("h1", null, "Meeting review")),
    h(
      "main",
      { class: "home" },
      h("h2", null, "Meetings"),
      h("p", { class: "lede" }, "Process a meeting to see what it changes in its deck, then decide each change."),
      h("ul", { class: "meetings" }, state.meetings.map(meetingRow)),
      state.meetings.length === 0 && h("p", { class: "muted" }, "No meetings found in evals/ or private/meetings/."),
    ),
  );
}

function meetingRow(m) {
  return h(
    "li",
    { class: "meeting", "data-meeting": m.id, "data-state": m.state.is },
    h(
      "div",
      { class: "meeting-main" },
      h("div", { class: "meeting-title" }, m.title),
      h("div", { class: "meeting-meta" }, m.date ?? "No date", " · ", m.origin === "evals" ? "Golden scenario" : "Private meeting"),
      h("div", { class: "chips" }, makerChip(m.maker), fontsChip(m.fonts)),
    ),
    h("div", { class: "meeting-side" }, stateBlock(m)),
  );
}

function makerChip(maker) {
  if (!maker) return h("span", { class: "chip muted" }, "Needs the maker, which is not built yet");
  return h("span", { class: `chip ${maker.simulated ? "simulated" : ""}`, title: maker.label }, maker.label);
}

function fontsChip(fonts) {
  if (!fonts.length) return null;
  return h("span", { class: "chip warn", title: fonts.map((f) => `${f.font} → ${f.family}`).join("\n") }, plural(fonts.length, "font substituted", "fonts substituted"));
}

function stateBlock(m) {
  const s = m.state;
  const processButton = (label) =>
    h("button", { class: "primary", "data-act": "process", disabled: !m.maker, onclick: () => processMeeting(m, false) }, label);
  switch (s.is) {
    case "processing":
      return h("div", { class: "progress" }, h("span", { class: "spinner" }), STEPS[s.step]);
    case "failed":
      return h(
        "div",
        { class: "failed" },
        h("p", { class: "error" }, s.message),
        s.problems.length > 0 && h("ul", { class: "problems" }, s.problems.map((p) => h("li", null, `${p.where}: ${p.message}`))),
        processButton("Process meeting"),
      );
    case "ready":
      return h(
        "div",
        { class: "ready" },
        h("span", { class: "count" }, `${s.decided} of ${s.total} decided`, s.applied ? " · applied" : ""),
        h("a", { class: "button primary", href: `#/m/${m.id}`, "data-act": "open" }, "Open review"),
        h("button", { class: "quiet", "data-act": "again", onclick: () => processMeeting(m, true) }, "Process again"),
      );
    default:
      return processButton("Process meeting");
  }
}

function renderMissing() {
  const row = state.missing;
  const why = row?.state ? `This meeting is ${row.state.is === "new" ? "not processed yet" : row.state.is}.` : row?.error ?? "No such meeting.";
  mount(
    root,
    h("header", { class: "top" }, h("a", { class: "back", href: "#/" }, "← Meetings")),
    h("main", { class: "home" }, h("p", null, why), h("a", { class: "button", href: "#/" }, "Back to the meetings")),
  );
}

function renderReview() {
  const v = state.view;
  const rail = M.deck(v.review, state.decisions);
  const entry = rail.find((e) => e.token === state.slide) ?? rail[0];
  const scroll = { rail: document.querySelector("#rail .rail-list")?.scrollTop, stage: document.getElementById("stage")?.scrollTop };
  const focus = document.activeElement?.matches?.("textarea[data-edit]") ? document.activeElement.dataset.edit : null;
  mount(
    root,
    h(
      "header",
      { class: "top" },
      h("a", { class: "back", href: "#/" }, "← Meetings"),
      h("div", { class: "title" }, h("strong", null, v.review.meeting.title), h("span", { class: "muted" }, v.review.meeting.date)),
      makerChip(v.maker),
    ),
    v.fonts.length > 0 &&
      h(
        "div",
        { id: "fonts", class: "banner" },
        "Fonts were substituted in the slide pictures, so line breaks may differ from PowerPoint: ",
        v.fonts.map((f) => `${f.font} → ${f.family}`).join(", "),
        ".",
      ),
    h(
      "div",
      { class: "review" },
      railPane(v, rail, entry),
      h(
        "main",
        { id: "stage", "data-view": entry.view, "data-slide": entry.token },
        needs(v, rail, entry),
        held(v, rail),
        heading(entry),
        entry.view === "text" ? textView(v, entry) : entry.view === "structural" ? structuralView(v, entry) : plainView(entry),
      ),
    ),
    bar(v),
  );
  if (scroll.rail !== undefined) document.querySelector("#rail .rail-list").scrollTop = scroll.rail;
  if (scroll.stage !== undefined) document.getElementById("stage").scrollTop = scroll.stage;
  if (focus) document.querySelector(`textarea[data-edit="${CSS.escape(focus)}"]`)?.focus();
  if (state.pinned !== null) showBubble(state.pinned);
}

function railPane(v, rail, entry) {
  const changed = rail.filter((e) => e.changes.length);
  return h(
    "aside",
    { id: "rail" },
    h(
      "div",
      { class: "rail-head" },
      h("span", { id: "rail-count" }, `${plural(v.review.changes.length, "change", "changes")} on ${plural(changed.length, "slide", "slides")}`),
      h(
        "div",
        { class: "toggle" },
        ["slides", "asks"].map((mode) =>
          h(
            "button",
            { "data-act": mode, "aria-pressed": String(state.rail === mode), onclick: () => ((state.rail = mode), renderReview()) },
            mode === "slides" ? "Slides" : "Asks",
          ),
        ),
      ),
    ),
    h("div", { class: "rail-list" }, state.rail === "slides" ? rail.map((e) => thumb(e, entry)) : asks(v, rail)),
  );
}

function thumb(e, entry) {
  const src = e.frame === "deleted" ? render("old", e.oldIndex) : render("new", e.newIndex);
  const done = e.changes.length > 0 && e.pending === 0;
  const tag = { added: "New", deleted: "Deleted", moved: `Moved from ${e.oldIndex}` }[e.frame];
  return h(
    "button",
    {
      class: `thumb ${e.token === entry.token ? "selected" : ""} ${e.frame ?? ""}`,
      "data-slide": e.token,
      "data-view": e.view,
      "data-count": e.changes.length,
      title: e.title,
      onclick: () => select(e.token),
    },
    h("span", { class: "thumb-number" }, e.newIndex ?? ""),
    h(
      "span",
      { class: "thumb-frame", style: { aspectRatio: `${state.view.slide.w} / ${state.view.slide.h}` } },
      img(src, "thumb"),
      e.changes.length > 0 && h("span", { class: `count ${done ? "done" : ""}` }, done ? "✓" : e.changes.length),
      e.flags.length > 0 && h("span", { class: "flagged", title: "A question names this slide" }, "?"),
      tag && h("span", { class: "tag" }, tag),
    ),
  );
}

function asks(v, rail) {
  return v.review.asks.map((ask) => {
    const changes = v.review.changes.filter((c) => c.ask_id === ask.id);
    return h(
      "div",
      { class: "ask" },
      h("p", null, ask.text),
      refs(ask.refs),
      h("div", { class: "chips" }, [...new Set(changes.map((c) => c.slide))].map((key) => slideChip(rail, key))),
    );
  });
}

function slideChip(rail, key) {
  const e = rail.find((x) => x.token === M.token(key));
  if (!e) return null;
  const where = e.frame === "deleted" ? `old slide ${e.oldIndex}` : e.frame === "added" ? `new slide ${e.newIndex}` : `Slide ${e.newIndex}`;
  return h("button", { class: "chip link", onclick: () => select(e.token) }, where[0].toUpperCase() + where.slice(1));
}

function needs(v, rail, entry) {
  const flags = [...v.review.flags].sort((a, b) => entry.flags.includes(b) - entry.flags.includes(a));
  return h(
    "section",
    { id: "needs", hidden: flags.length === 0 },
    h("h3", null, "Needs you"),
    flags.map((f) =>
      h(
        "div",
        { class: `flag ${entry.flags.includes(f) ? "here" : ""}`, "data-flag": f.id },
        h("p", { class: "question" }, f.question),
        refs(f.refs),
        h("div", { class: "chips" }, f.slides.map((s) => slideChip(rail, s))),
      ),
    ),
  );
}

function held(v, rail) {
  return h(
    "details",
    { id: "held", hidden: v.review.held.length === 0 },
    h("summary", null, `Heard, not changed (${v.review.held.length})`),
    v.review.held.map((x) =>
      h("div", { class: "held", "data-held": x.id }, h("p", null, x.text), refs(x.refs), h("div", { class: "chips" }, x.slides.map((s) => slideChip(rail, s)))),
    ),
  );
}

function heading(e) {
  const where = {
    added: `New slide ${e.newIndex}`,
    deleted: `Slide ${e.oldIndex} of the old deck, deleted`,
    moved: `Slide ${e.newIndex}, moved from slide ${e.oldIndex} to slide ${e.newIndex}`,
  }[e.frame] ?? `Slide ${e.newIndex}`;
  return h("div", { class: "slide-heading" }, h("span", { class: "where" }, where), h("span", { class: "slide-title" }, e.title));
}

function bar(v) {
  const { decided, total, pending } = M.tally(v.review, state.decisions);
  const f = state.final;
  return h(
    "footer",
    { id: "bar" },
    h("span", { id: "decided" }, `${decided} of ${total} decided`),
    h(
      "div",
      { class: "outcome" },
      state.notice && h("span", { class: "error" }, state.notice),
      f &&
        h(
          "div",
          null,
          h("span", { id: "final-path" }, `Final deck: ${f.path}`),
          h("a", { id: "download", href: f.download, download: "" }, "Download"),
          h("span", { class: "muted" }, `${f.kept_new.length} kept new, ${f.edited.length} edited, ${f.kept_old.length} kept old, ${f.dropped.length} dropped`),
        ),
      !f && state.stale && h("span", { class: "stale" }, "A decision changed, so the final deck is out of date. Apply again."),
      state.published && h("span", { id: "publish-note", class: "muted" }, "Nothing was uploaded. Publishing needs a OneDrive connection that IT has not approved yet."),
    ),
    h(
      "div",
      { class: "actions" },
      h(
        "button",
        { class: "stub", "data-act": "publish", onclick: () => ((state.published = true), renderReview()) },
        "Publish to OneDrive",
        h("small", null, "Not connected: needs IT approval"),
      ),
      h("button", { class: "primary", "data-act": "apply", disabled: pending > 0, onclick: applyDecisions }, pending > 0 ? `Decide ${pending} more` : "Apply decisions"),
    ),
  );
}

async function applyDecisions() {
  const { status, body } = await call("POST", `${meetingUrl(state.view.id)}/apply`, {});
  if (status === 200) {
    state.final = body;
    state.stale = false;
    state.notice = null;
  } else {
    state.notice = body.error ?? `Apply failed (${status}).`;
  }
  renderReview();
}

document.addEventListener("keydown", (event) => {
  if (!state.view || parseHash().mid === null) return;
  if (event.key === "Escape") {
    if (state.editing !== null) return cancelEdit({ id: state.editing });
    return unpin();
  }
  if (event.target instanceof Element && event.target.matches("textarea, input")) return;
  if (event.key !== "ArrowDown" && event.key !== "ArrowUp") return;
  event.preventDefault();
  const rail = M.deck(state.view.review, state.decisions);
  const at = rail.findIndex((e) => e.token === state.slide);
  const next = rail[Math.max(0, Math.min(rail.length - 1, at + (event.key === "ArrowDown" ? 1 : -1)))];
  if (next && next.token !== state.slide) select(next.token);
});

document.addEventListener("click", (event) => {
  if (state.pinned === null || state.editing !== null) return;
  if (event.target.closest("#bubble, .mark")) return;
  unpin();
});

addEventListener("resize", placeBubble);

document.addEventListener("scroll", placeBubble, true);

ui.rerender = renderReview;
addEventListener("hashchange", route);

route();
