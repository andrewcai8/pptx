export function h(tag, attrs, ...children) {
  const el = document.createElement(tag);
  for (const [key, value] of Object.entries(attrs ?? {})) {
    if (value === null || value === undefined || value === false) continue;
    if (key.startsWith("on")) el.addEventListener(key.slice(2), value);
    else if (key === "style") Object.assign(el.style, value);
    else if (key === "value") el.value = value;
    else el.setAttribute(key, value === true ? "" : String(value));
  }
  const kids = children.flat(Infinity).filter((c) => c !== null && c !== undefined && c !== false);
  el.append(...kids.map((c) => (c instanceof Node ? c : String(c))));
  return el;
}
