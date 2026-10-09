import { h } from "./dom.js";
import { serial } from "./model.js";

export const state = {
  meetings: [],
  mine: new Set(),
  poll: null,
  view: null,
  missing: null,
  decisions: {},
  final: null,
  stale: false,
  slide: null,
  rail: "slides",
  pinned: null,
  open: null,
  editing: null,
  drafts: {},
  errors: {},
  notice: null,
  offline: null,
  published: false,
};

export const mount = (el, ...children) => el.replaceChildren(...children.flat().filter((c) => c instanceof Node));

const images = new Map();

export async function call(method, url, body) {
  const init = { method, headers: { "Content-Type": "application/json" } };
  if (body !== undefined) init.body = JSON.stringify(body);
  let response;
  try {
    response = await fetch(url, init);
  } catch {
    return { status: 0, body: { error: "The review server did not answer. Check that it is still running, then try again." } };
  }
  try {
    return { status: response.status, body: await response.json() };
  } catch {
    return { status: 0, body: { error: `The review server answered ${response.status} with something other than JSON.` } };
  }
}

export const post = serial();

export const meetingUrl = (id) => `/api/meetings/${id}`;

export function img(url, role) {
  const key = `${role} ${url}`;
  if (!images.has(key)) images.set(key, h("img", { src: url, alt: "", draggable: "false" }));
  return images.get(key);
}

export const render = (side, n) => state.view.images[side].replace("{n}", String(n));

export const ui = { rerender: () => {} };
