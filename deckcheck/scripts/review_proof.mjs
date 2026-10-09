import { spawn, spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import { get } from "node:http";
import { chmodSync, copyFileSync, existsSync, mkdirSync, readFileSync, renameSync, rmSync, writeFileSync } from "node:fs";
import { join, resolve } from "node:path";

const OUT = process.argv[2] ?? "artifacts/review-proof";
const PORT = 8790 + Math.floor(Math.random() * 100);
const CDP = PORT + 1000;
const BASE = `http://127.0.0.1:${PORT}`;
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const sha = (path) => createHash("sha256").update(readFileSync(path)).digest("hex");
let failures = 0;
const check = (ok, msg) => {
  console.log(`${ok ? "ok " : "BAD"} ${msg}`);
  if (!ok) failures++;
};

const SCENARIOS = ["insurance-workshop-prep", "solar-market-refresh", "fmcg-diagnostic-timeline"];
const PRIVATE = "private/meetings/my meeting";
const PRIVATE_ID = "private/my%20meeting";
const PRIVATE_WORK = "artifacts/review/private/my meeting";
const MAKER = "uv run --project deckcheck meeting process {meeting} --out {dir} --shareable";
const CONSENT = "This sends the meeting notes, data and deck to Claude (Anthropic). Only continue if your firm allows sharing this deck with an AI service.";
const SOLAR = JSON.parse(readFileSync("evals/solar-market-refresh/changeset.json", "utf8"));
if (existsSync(PRIVATE)) {
  console.log(`REVIEW PROOF FAIL (${PRIVATE} already exists; this proof creates and deletes it, so move it away first)`);
  process.exit(1);
}
rmSync(OUT, { recursive: true, force: true });
mkdirSync(join(OUT, "bin"), { recursive: true });
rmSync(PRIVATE_WORK, { recursive: true, force: true });
mkdirSync(PRIVATE, { recursive: true });
copyFileSync(SOLAR.source.path, join(PRIVATE, "before.pptx"));
writeFileSync(join(PRIVATE, "notes.md"), "# Solar update call\n\nDate: 2026-10-08\n\nAna Ruiz: The market is now $400m, not $360m. Keep the growth rate as it is.\n");
const fakeClaude = join(OUT, "bin", "claude");
writeFileSync(fakeClaude, `#!/usr/bin/env node
const fs = require("node:fs");
if (process.argv[2] === "auth") { console.log('{"loggedIn": true}'); process.exit(0); }
const cs = JSON.parse(fs.readFileSync(${JSON.stringify(resolve("evals/solar-market-refresh/changeset.json"))}, "utf8"));
cs.source.path = "before.pptx";
fs.writeFileSync("changeset.json", JSON.stringify(cs, null, 2));
console.log('{"type": "result", "result": "fake claude wrote changeset.json"}');
`);
chmodSync(fakeClaude, 0o755);

const server = spawn("uv", ["run", "--quiet", "--project", "deckcheck", "review", "serve", "--port", String(PORT), "--maker", MAKER], {
  stdio: ["ignore", "pipe", "pipe"],
  env: { ...process.env, PATH: `${resolve(OUT, "bin")}:${process.env.PATH}` },
});
let serverLog = "";
server.stderr.on("data", (d) => (serverLog += d));
const chrome = spawn(process.env.CHROME ?? "/usr/bin/google-chrome", ["--headless=new", "--no-sandbox", "--disable-gpu", "--hide-scrollbars", `--remote-debugging-port=${CDP}`,
  `--user-data-dir=/tmp/review-proof-chrome-${PORT}`, "--window-size=1440,900", "about:blank"], { stdio: "ignore" });
const stop = () => {
  server.kill();
  chrome.kill();
  rmSync(PRIVATE, { recursive: true, force: true });
  rmSync(PRIVATE_WORK, { recursive: true, force: true });
};
process.on("exit", stop);
const abort = (why) => {
  console.log(`REVIEW PROOF FAIL (${why})`);
  process.exit(1);
};

for (const t0 = Date.now(); ; await sleep(250)) {
  if (server.exitCode !== null || server.signalCode !== null) abort(`the review server exited early: ${serverLog.trim()}`);
  if (await fetch(`${BASE}/api/meetings`).then((r) => r.ok, () => false)) break;
  if (Date.now() - t0 > 30000) abort(`the review server did not answer in 30 s: ${serverLog.trim()}`);
}
SCENARIOS.forEach((name) => rmSync(`artifacts/review/evals/${name}`, { recursive: true, force: true }));

async function until(probe, what, ms = 30000) {
  const t0 = Date.now();
  for (;;) {
    try { const v = await probe(); if (v) return v; } catch {}
    if (Date.now() - t0 > ms) throw new Error(`timed out waiting for ${what}`);
    await sleep(250);
  }
}

const target = await until(async () => (await (await fetch(`http://127.0.0.1:${CDP}/json`)).json()).find((t) => t.type === "page"), "chrome");
const ws = new WebSocket(target.webSocketDebuggerUrl);
await new Promise((r) => ws.addEventListener("open", r));
let seq = 0;
const waiting = new Map();
const dialogs = { accept: true, seen: [] };
ws.addEventListener("message", (e) => {
  const m = JSON.parse(e.data);
  if (m.id && waiting.has(m.id)) { waiting.get(m.id)(m); waiting.delete(m.id); }
  if (m.method === "Page.javascriptDialogOpening") {
    dialogs.seen.push(m.params.message);
    send("Page.handleJavaScriptDialog", { accept: dialogs.accept });
  }
});
const send = (method, params = {}) => new Promise((r) => { const id = ++seq; waiting.set(id, r); ws.send(JSON.stringify({ id, method, params })); });
const js = async (expr) => {
  const r = await send("Runtime.evaluate", { expression: expr, awaitPromise: true, returnByValue: true });
  if (r.result.exceptionDetails) throw new Error(`${expr}: ${r.result.exceptionDetails.exception?.description}`);
  return r.result.result.value;
};
const q = (sel) => JSON.stringify(sel);
const exists = (sel) => js(`!!document.querySelector(${q(sel)})`);
const text = (sel) => js(`document.querySelector(${q(sel)})?.textContent ?? null`);
const attr = (sel, name) => js(`document.querySelector(${q(sel)})?.getAttribute(${q(name)}) ?? null`);
const waitFor = (sel, ms) => until(() => exists(sel), sel, ms);
const shot = async (name) => {
  await sleep(300);
  const r = await send("Page.captureScreenshot", { format: "png" });
  writeFileSync(join(OUT, `${name}.png`), Buffer.from(r.result.data, "base64"));
  console.log(`    shot ${join(OUT, `${name}.png`)}`);
};
async function center(sel) {
  await js(`document.querySelector(${q(sel)}).scrollIntoView({ block: "nearest" })`);
  await sleep(100);
  return js(`(() => { const r = document.querySelector(${q(sel)}).getBoundingClientRect(); return [r.left + r.width / 2, r.top + r.height / 2]; })()`);
}
async function hover(sel) {
  const [x, y] = await center(sel);
  await send("Input.dispatchMouseEvent", { type: "mouseMoved", x, y });
  await sleep(300);
}
async function click(sel) {
  await waitFor(sel);
  const [x, y] = await center(sel);
  for (const type of ["mouseMoved", "mousePressed", "mouseReleased"]) await send("Input.dispatchMouseEvent", { type, x, y, button: "left", clickCount: 1 });
  await sleep(400);
}
async function typeInto(sel, value) {
  await click(sel);
  await js(`(() => { const t = document.querySelector(${q(sel)}); t.select(); })()`);
  for (const type of ["keyDown", "keyUp"]) await send("Input.dispatchKeyEvent", { type, key: "Backspace", code: "Backspace", windowsVirtualKeyCode: 8 });
  if (value) await send("Input.insertText", { text: value });
  await sleep(200);
}
const go = async (hash) => {
  await send("Page.navigate", { url: `${BASE}/${hash}` });
  await sleep(500);
};
const mark = (cid) => `#stage [data-side="new"] [data-mark][data-changes~=${q(cid)}], #stage [data-mark][data-changes~=${q(cid)}]`;
const block = (cid) => `[data-change=${q(cid)}]`;

async function decide(cid, how, edited) {
  const scope = block(cid);
  if (how === "edited") {
    await click(`${scope} [data-decide="edit"]`);
    await typeInto(`${scope} textarea[data-edit]`, edited);
    await click(`${scope} [data-act="save"]`);
  } else {
    await click(`${scope} [data-decide="${how}"]`);
  }
  const want = how === "edited" ? "edited" : how;
  await until(async () => (await attr(scope, "data-decision")) === want, `${cid} shows ${want}`, 5000).catch(() => {});
  check((await attr(scope, "data-decision")) === want, `${cid}: the UI shows ${want}`);
}

async function processThroughHome(name) {
  await go("#/");
  const row = `[data-meeting="evals/${name}"]`;
  await waitFor(row);
  check((await attr(row, "data-state")) === "new", `${name} starts unprocessed`);
  await click(`${row} [data-act="process"]`);
  await until(async () => (await js("location.hash")).startsWith(`#/m/evals/${name}`) && exists("#rail [data-slide]"), `${name} review opens`, 240000);
  await waitFor("#stage[data-view]");
}

async function selectSlide(token, view) {
  await click(`#rail [data-slide=${q(token)}]`);
  await until(async () => (await attr("#stage", "data-slide")) === token, `slide ${token} selected`, 5000);
  const got = await attr("#stage", "data-view");
  check(got === view, `slide ${token} routes to the ${view} view (got ${got})`);
}

function proveApply(name, decisions) {
  const dir = join(OUT, name);
  mkdirSync(dir, { recursive: true });
  const cs = JSON.parse(readFileSync(`evals/${name}/changeset.json`, "utf8"));
  for (const c of cs.changes) c.decision = decisions[c.id];
  writeFileSync(join(dir, "changeset.json"), JSON.stringify(cs, null, 2));
  const cli = spawnSync("uv", ["run", "--quiet", "--project", "deckcheck", "changeset", "apply", join(dir, "changeset.json"), "--out", join(dir, "cli.pptx")], { encoding: "utf8" });
  check(cli.status === 0, `${name}: changeset apply on the committed ChangeSet with the same decisions exits 0 (${cli.stdout.trim()})`);
  const app = JSON.parse(readFileSync(`artifacts/review/evals/${name}/changeset.json`, "utf8"));
  const written = Object.fromEntries(app.changes.map((c) => [c.id, c.decision ?? "pending"]));
  check(JSON.stringify(written) === JSON.stringify(decisions), `${name}: the server wrote exactly the decisions set in the UI ${JSON.stringify(written)}`);
  const ui = sha(`artifacts/review/evals/${name}/final.pptx`), oracle = sha(join(dir, "cli.pptx"));
  check(ui === oracle, `${name}: final.pptx from the UI is byte-identical to changeset apply (${ui.slice(0, 12)} vs ${oracle.slice(0, 12)})`);
  const render = spawnSync("uv", ["run", "--quiet", "--project", "deckcheck", "deckcheck", "render", `artifacts/review/evals/${name}/final.pptx`, "--out", join(dir, "final-render")], { encoding: "utf8" });
  check(render.status === 0, `${name}: the final deck renders to ${join(dir, "final-render")}`);
}

async function applyThroughBar(name) {
  check(/^\s*(\d+) of \1 decided/.test(await text("#decided")), `${name}: bar says every change is decided (${await text("#decided")})`);
  await click('#bar [data-act="apply"]');
  await until(async () => ((await text("#final-path")) ?? "").includes(`artifacts/review/evals/${name}/final.pptx`), `${name} final path`, 60000);
  const href = await attr("a#download", "href");
  const res = await fetch(new URL(href, BASE));
  const body = Buffer.from(await res.arrayBuffer());
  check(res.ok && createHash("sha256").update(body).digest("hex") === sha(`artifacts/review/evals/${name}/final.pptx`), `${name}: the download link serves the final deck`);
}

try {
  await send("Page.enable");
  await send("Emulation.setDeviceMetricsOverride", { width: 1440, height: 900, deviceScaleFactor: 1, mobile: false });

  await go("#/");
  await waitFor("[data-meeting]");
  const rows = await js(`[...document.querySelectorAll("[data-meeting]")].map((r) => r.dataset.meeting)`);
  check(["fmcg-diagnostic-timeline", "insurance-workshop-prep", "rcc-flexibility-wording", "retail-impact-title", "solar-market-refresh"].every((n) => rows.includes(`evals/${n}`)), `home lists the five golden meetings (${rows.join(", ")})`);
  check(/simulated/i.test(await text('[data-meeting="evals/solar-market-refresh"]')), "home labels the golden maker as simulated");
  await shot("01-home");

  const ins = "insurance-workshop-prep";
  await processThroughHome(ins);
  const railCount = await js(`document.querySelectorAll("#rail [data-slide]").length`);
  check(railCount === 20, `insurance rail holds 19 new slides and 1 deleted ghost (got ${railCount})`);
  check(/4 changes on 3 slides/.test(await text("#rail-count")), `rail header counts the changes (${await text("#rail-count")})`);
  check((await attr('#rail [data-slide="s301"]', "data-count")) === "2", "rail marks slide 301 with 2 changes");
  await shot("02-rail");

  await selectSlide("s301", "text");
  await hover(mark("headline-share"));
  check(!(await js(`document.querySelector("#bubble").hidden`)), "hovering a highlight opens the bubble");
  const bubble = await text("#bubble");
  check(["ca. 50%", "ca. 48%", "00:06:24", "Mojca Zupan", "Use the file", "survey's 2027 projection"].every((s) => bubble.includes(s)), "the bubble shows before, after, the transcript line and the rationale");
  const [bx] = await js(`(() => { const r = document.querySelector("#bubble").getBoundingClientRect(); return [r.left]; })()`);
  check(bx > 1440 / 2, `the bubble sits toward the right of the screen (left edge ${Math.round(bx)}px)`);
  await shot("03-text-bubble");
  await decide("headline-share", "keep_new");

  await click(mark("survey-footnote"));
  await click(`${block("survey-footnote")} [data-decide="edit"]`);
  await typeInto(`${block("survey-footnote")} textarea[data-edit]`, "3. Insurers' Association member survey 2026");
  await shot("06-edit-myself");
  await click(`${block("survey-footnote")} [data-act="save"]`);
  await until(async () => (await attr(block("survey-footnote"), "data-decision")) === "edited", "survey-footnote edited", 5000).catch(() => {});
  check((await attr(block("survey-footnote"), "data-decision")) === "edited", "survey-footnote: the UI shows edited");

  await click(`${block("survey-footnote")} [data-decide="edit"]`);
  await typeInto(`${block("survey-footnote")} textarea[data-edit]`, "");
  await click(`${block("survey-footnote")} [data-act="save"]`);
  await until(() => exists(`${block("survey-footnote")} [data-error]`), "an error for an empty paragraph", 5000).catch(() => {});
  check(!!(await text(`${block("survey-footnote")} [data-error]`)), `an edit the engine refuses shows its message (${await text(`${block("survey-footnote")} [data-error]`)})`);
  await click(`${block("survey-footnote")} [data-act="cancel"]`);
  check((await attr(block("survey-footnote"), "data-decision")) === "edited", "a refused edit leaves the earlier decision in place");
  await send("Input.dispatchKeyEvent", { type: "keyDown", key: "Escape", code: "Escape", windowsVirtualKeyCode: 27 });

  await js(`document.querySelector("#needs").scrollIntoView({ block: "start" })`);
  check(await exists('#needs [data-flag="which-regulator-slide"]'), "Needs you shows the regulator flag");
  check(/00:08:05/.test(await text("#needs")), "the flag carries its transcript refs");
  check(!(await js(`document.querySelector("details#held").open`)), "heard, not changed starts collapsed");
  check(/Heard, not changed/i.test(await text("details#held summary")), "the held list is labelled");
  await shot("05-flags");

  await selectSlide("s326", "structural");
  check(/deleted/i.test(await text('#stage [data-side="new"]')), "a deleted slide shows Deleted on the new side");
  await shot("04b-structural-deleted");
  await decide("delete-credentials", "keep_old");
  await selectSlide("s322", "structural");
  check(/moved from slide 14/i.test(await text("#stage")), "a moved slide says where it moved from");
  await decide("innovation-after-overview", "keep_new");
  await applyThroughBar(ins);
  await shot("07-post-apply");
  proveApply(ins, { "headline-share": "keep_new", "survey-footnote": { edited: "3. Insurers' Association member survey 2026" }, "delete-credentials": "keep_old", "innovation-after-overview": "keep_new" });

  const sol = "solar-market-refresh";
  await processThroughHome(sol);
  await selectSlide("s2147478638", "structural");
  const sides = await js(`["old", "new"].map((s) => document.querySelectorAll('#stage [data-side="' + s + '"] [data-mark]').length)`);
  check(sides[0] >= 5 && sides[1] >= 5, `solar slide 10 highlights every changed shape on both sides (${sides})`);
  check(await js(`document.querySelectorAll("#changes [data-change]").length`) === 6, "solar slide 10 lists its 6 changes");
  await hover(`#changes ${block("chart-2022-bar")}`);
  await shot("04-structural-solar");
  await decide("title-market-size", "keep_new");
  await decide("title-cagr", "keep_old");
  await decide("header-market-size", "edited", "$410m market set to grow");
  await decide("table-market-size", "keep_new");
  await decide("chart-2022-bar", "edited", "405");
  await decide("cagr-label", "keep_old");
  await selectSlide("s2147480153", "text");
  await click(mark("contents-cagr"));
  await decide("contents-cagr", "keep_new");
  await applyThroughBar(sol);
  await click('#rail [data-act="asks"]');
  check(/market size/i.test((await text("#rail")) ?? ""), "the rail groups the changes by ask on request");
  await shot("08-asks");
  await click('#rail [data-act="slides"]');
  proveApply(sol, { "title-market-size": "keep_new", "title-cagr": "keep_old", "header-market-size": { edited: "$410m market set to grow" }, "table-market-size": "keep_new", "chart-2022-bar": { edited: "405" }, "cagr-label": "keep_old", "contents-cagr": "keep_new" });

  const fm = "fmcg-diagnostic-timeline";
  await processThroughHome(fm);
  await selectSlide("a-diagnostic-slide", "structural");
  check(/new slide/i.test(await text('#stage [data-side="old"]')), "an added slide shows an empty New slide panel on the old side");
  await shot("04c-structural-added");
  await decide("diagnostic-slide", "keep_new");
  await decide("diagnostic-title", "edited", "STEP 1: Diagnose in 6 weeks");
  await decide("diagnostic-bullets", "keep_new");
  await applyThroughBar(fm);
  proveApply(fm, { "diagnostic-slide": "keep_new", "diagnostic-title": { edited: "STEP 1: Diagnose in 6 weeks" }, "diagnostic-bullets": "keep_new" });

  await click('#bar [data-act="publish"]');
  check(/IT/.test((await text("#publish-note")) ?? ""), "publish is a labelled stub");

  await go("#/");
  const row = `[data-meeting=${q(PRIVATE_ID)}]`;
  await waitFor(row);
  check(/Solar update call/.test(await text(row)) && /Maker: uv run/.test(await text(row)), `a notes-only meeting in a folder with a space is listed with its maker (${await text(row)})`);
  await shot("09-private-listed");
  dialogs.accept = false;
  await click(`${row} [data-act="process"]`);
  await sleep(1500);
  check(dialogs.seen.at(-1) === CONSENT, `Process meeting asks for consent first (${dialogs.seen.at(-1)})`);
  check((await attr(row, "data-state")) === "new" && !existsSync(`${PRIVATE_WORK}.partial`) && !existsSync(PRIVATE_WORK), "cancel sends nothing and the meeting stays unprocessed");
  await shot("10-consent-cancelled");
  dialogs.accept = true;
  await click(`${row} [data-act="process"]`);
  check(dialogs.seen.length === 2 && dialogs.seen[1] === CONSENT, "confirm asks once more on the second click");
  await until(async () => (await js("location.hash")).startsWith(`#/m/${PRIVATE_ID}`) && exists("#rail [data-slide]"), "the private review opens", 240000);
  check(/fake claude wrote changeset\.json/.test(readFileSync(join(PRIVATE_WORK, "claude.json"), "utf8")), "meeting process ran the fake claude in its staged folder");
  await selectSlide("s2147478638", "structural");
  await shot("11-private-review");

  renameSync(join(PRIVATE, "before.pptx"), join(OUT, "before.pptx.away"));
  await click(`${block("title-market-size")} [data-decide="keep_new"]`);
  await until(() => exists(`${block("title-market-size")} [data-error]`), "a message for a refused click", 5000).catch(() => {});
  const refused = await text(`${block("title-market-size")} [data-error]`);
  check(/cannot read .*before\.pptx/.test(refused ?? ""), `a Keep click the engine refuses shows its message (${refused})`);
  check((await attr(block("title-market-size"), "data-decision")) === "pending", "a refused click leaves the change undecided");
  await shot("12-refused-click");
  renameSync(join(OUT, "before.pptx.away"), join(PRIVATE, "before.pptx"));
  await decide("title-market-size", "keep_new");
  check(!(await exists(`${block("title-market-size")} [data-error]`)), "the message clears once the click succeeds");
  await decide("title-cagr", "keep_old");
  await decide("header-market-size", "edited", "$400m market set to grow");
  await decide("table-market-size", "keep_new");
  await decide("chart-2022-bar", "keep_new");
  await decide("cagr-label", "keep_old");
  await selectSlide("s2147480153", "text");
  await click(mark("contents-cagr"));
  await decide("contents-cagr", "keep_new");
  await click('#bar [data-act="apply"]');
  await until(async () => ((await text("#final-path")) ?? "").includes(`${PRIVATE_WORK}/final.pptx`), "private final path", 60000);
  await shot("13-private-applied");
  copyFileSync(join(PRIVATE_WORK, "changeset.json"), join(OUT, "private-changeset.json"));
  const cli = spawnSync("uv", ["run", "--quiet", "--project", "deckcheck", "changeset", "apply", join(OUT, "private-changeset.json"), "--out", join(OUT, "private-cli.pptx")], { encoding: "utf8" });
  check(cli.status === 0, `changeset apply on the private review's ChangeSet exits 0 (${cli.stdout.trim()})`);
  check(sha(join(PRIVATE_WORK, "final.pptx")) === sha(join(OUT, "private-cli.pptx")), "the private final.pptx is byte-identical to changeset apply");
  const res = await fetch(new URL(await attr("a#download", "href"), BASE));
  check(res.ok && res.headers.get("content-disposition") === 'attachment; filename="my-meeting-final.pptx"', `the download is named my-meeting-final.pptx (${res.headers.get("content-disposition")})`);
  for (const probe of ["private/..", "private/%2E%2E", "private/..%2Fmeetings%2Fmy%20meeting", "private/%2Fetc%2Fpasswd", "private/my%20meeting.partial", "evals/solar-market-refresh.partial", "evals/..%2F..%2Fdeckcheck"]) {
    const status = await new Promise((done, fail) => get({ host: "127.0.0.1", port: PORT, path: `/api/meetings/${probe}` }, (r) => (r.resume(), done(r.statusCode))).on("error", fail));
    check(status === 404, `${probe} sent verbatim answers 404 (got ${status})`);
  }
} catch (e) {
  check(false, `drive crashed: ${e.message}`);
  await shot("crash").catch(() => {});
}
ws.close();
stop();
console.log(failures ? `REVIEW PROOF FAIL (${failures} checks)` : "REVIEW PROOF PASS");
console.log(`evidence: ${OUT}`);
process.exit(failures ? 1 : 0);
