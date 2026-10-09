from __future__ import annotations

import copy
from collections.abc import Sequence
from dataclasses import dataclass

from lxml import etree
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement

from deckcheck.changeset.model import Paragraph

A_R, A_FLD, A_BR, A_T = qn("a:r"), qn("a:fld"), qn("a:br"), qn("a:t")
LINE_BREAK = "\v"


@dataclass(frozen=True)
class Atom:
    """One content child of a paragraph and the characters it holds in the paragraph's text."""

    el: etree._Element
    lo: int
    hi: int

    @property
    def is_run(self) -> bool:
        return self.el.tag == A_R


@dataclass(frozen=True)
class Splice:
    """The smallest changing range of a quote and its new text, in the paragraph's source coordinates, so new
    characters land in the run that held the characters they replace and keep its formatting."""

    start: int
    end: int
    new: str


class SpliceError(ValueError):
    pass


def atoms(p: etree._Element) -> list[Atom]:
    out, at = [], 0
    for child in p:
        if child.tag in (A_R, A_FLD):
            text = child.findtext(A_T) or ""
        elif child.tag == A_BR:
            text = LINE_BREAK
        else:
            continue
        out.append(Atom(child, at, at + len(text)))
        at += len(text)
    return out


def paragraph_text(p: etree._Element) -> str:
    return "".join((a.el.findtext(A_T) or "") if a.el.tag != A_BR else LINE_BREAK for a in atoms(p))


def clip(s: str, n: int = 120) -> str:
    return s if len(s) <= n else s[: n - 1] + "…"


def occurrences(text: str, quote: str) -> list[int]:
    return [i for i in range(len(text) - len(quote) + 1) if text.startswith(quote, i)]


# think-cell writes a label as a datetime field whose custom format is the label's characters, each quoted
# with '' padding. Measured on the corpus: in 1568 of 1568 datetime fields every format character is quoted and
# the quoted characters spell the field text. A field that does not is a real date or a slide number, which no
# edit may rewrite.
def decode(kind: str) -> str | None:
    """The text a datetime field's format spells, or None when a character sits outside the quotes, where
    PowerPoint reads it as a date code."""
    if not kind.startswith("datetime"):
        return None
    out, quoted = [], False
    for c in kind[8:]:
        if c == "'":
            quoted = not quoted
        elif quoted:
            out.append(c)
        else:
            return None
    return None if quoted else "".join(out)


def is_literal_field(el: etree._Element) -> bool:
    return el.tag == A_FLD and decode(el.get("type", "")) == (el.findtext(A_T) or "")


def trim(old: str, new: str) -> tuple[int, int, str]:
    pre = 0
    while pre < min(len(old), len(new)) and old[pre] == new[pre]:
        pre += 1
    suf = 0
    while suf < min(len(old), len(new)) - pre and old[-1 - suf] == new[-1 - suf]:
        suf += 1
    return pre, len(old) - suf, new[pre : len(new) - suf]


def plan_splice(p: etree._Element, at: int, old: str, new: str) -> Splice:
    a, b, core = trim(old, new)
    s = Splice(at + a, at + b, core)
    if LINE_BREAK in core or (LINE_BREAK in old[a:b]):
        raise SpliceError("adding or removing a line break (\\v) is not supported; keep the line breaks where they are")
    xs = atoms(p)
    writes = _writes(xs, s)
    if s.start == s.end and core and xs and not writes:
        raise SpliceError(f"no run holds character {s.start} to write into; quote text from a run")
    fields = [w.atom.el for w in writes if w.atom.el.tag == A_FLD]
    for el in fields:
        if not is_literal_field(el):
            raise SpliceError(f"the change touches a {clip(el.get('type', 'field'), 40)!r} field, which only PowerPoint fills in")
    if fields and len(writes) > 1:
        raise SpliceError("the change crosses the edge of a field; change the field's text or the run's text, not both")
    for w in writes:
        if w.atom.el.tag == A_FLD:
            _field_type(w.atom.el, w)
    return s


@dataclass(frozen=True)
class Write:
    """What one atom gets: its characters [a, b) replaced by `piece`."""

    atom: Atom
    a: int
    b: int
    piece: str


def _writes(xs: Sequence[Atom], s: Splice) -> list[Write]:
    """The atoms a splice writes. An insertion goes into the run or literal field holding the character before
    it; a replacement into every atom it touches, with the new text in the first. plan_splice checks these
    writes and write_splice makes them, so the two cannot disagree."""
    if s.start == s.end:
        target = _insertion_target(xs, s.start)
        return [Write(target, s.start - target.lo, s.start - target.lo, s.new)] if target and s.new else []
    touched = [x for x in xs if x.lo < s.end and x.hi > s.start]
    return [Write(x, max(s.start - x.lo, 0), min(s.end, x.hi) - x.lo, s.new if i == 0 else "") for i, x in enumerate(touched)]


def _insertion_target(xs: Sequence[Atom], at: int) -> Atom | None:
    writable = [x for x in xs if x.is_run or is_literal_field(x.el)]
    before = next((x for x in writable if x.lo < at <= x.hi), None)
    return before or next((x for x in writable if x.lo == at), None)


def _field_type(el: etree._Element, w: Write) -> str:
    if "'" in w.piece:
        raise SpliceError("a field's text cannot take an apostrophe, which its format uses for quoting")
    old = el.findtext(A_T) or ""
    kind = retype(el.get("type"), w.a, w.b, w.piece)
    if decode(kind) != old[: w.a] + w.piece + old[w.b :]:
        raise SpliceError(f"the field's format {clip(kind, 40)!r} would not spell its new text; change the label in think-cell")
    return kind


def write_splice(p: etree._Element, s: Splice) -> None:
    writes = _writes(atoms(p), s)
    if not writes and s.new:
        _new_run(p, s.new)
    for w in writes:
        t = w.atom.el.find(A_T)
        if w.atom.el.tag == A_FLD:
            w.atom.el.set("type", _field_type(w.atom.el, w))
        t.text = (t.text or "")[: w.a] + w.piece + (t.text or "")[w.b :]
    for w in writes:
        if w.atom.is_run and not (w.atom.el.findtext(A_T) or "") and len(p.findall(A_R)) > 1:
            p.remove(w.atom.el)


def retype(kind: str, a: int, b: int, piece: str) -> str:
    """Rewrite a literal datetime field's format so its characters spell the new text, keeping the quote
    skeleton: substitute in place when the length holds, else insert or delete next to a neighbour. A field
    with no characters has no quote to write into, so its new characters get one of their own."""
    head, body = kind[:8], kind[8:]
    pos = [i for i, c in enumerate(body) if c != "'"]
    if not pos:
        return f"{kind}'{piece}'"
    if b - a == len(piece):
        chars = list(body)
        for k, c in enumerate(piece):
            chars[pos[a + k]] = c
        return head + "".join(chars)
    if a < b:
        lo, hi = pos[a], pos[b - 1] + 1
        return head + body[:lo] + piece + "".join(c for c in body[lo:hi] if c == "'") + body[hi:]
    at = pos[a - 1] + 1 if a > 0 else pos[0]
    return head + body[:at] + piece + body[at:]


def _new_run(p: etree._Element, text: str) -> None:
    end = p.find(qn("a:endParaRPr"))
    run = _run(text, end)
    if end is not None:
        end.addprevious(run)
    else:
        p.append(run)


def _run(text: str, props: etree._Element | None) -> etree._Element:
    run = OxmlElement("a:r")
    if props is not None:
        rpr = copy.deepcopy(props)
        rpr.tag = qn("a:rPr")
        run.append(rpr)
    t = OxmlElement("a:t")
    t.text = text
    run.append(t)
    return run


def _runs(text: str, props: etree._Element | None) -> list[etree._Element]:
    out: list[etree._Element] = []
    for i, line in enumerate(text.split(LINE_BREAK)):
        if i:
            br = OxmlElement("a:br")
            if props is not None:
                br.append(copy.deepcopy(props))
                br[0].tag = qn("a:rPr")
            out.append(br)
        out.append(_run(line, props))
    return out


def new_paragraph(anchor: etree._Element, text: str) -> etree._Element:
    """A paragraph styled like `anchor`: its paragraph properties and its first run's properties."""
    p = OxmlElement("a:p")
    ppr = anchor.find(qn("a:pPr"))
    if ppr is not None:
        p.append(copy.deepcopy(ppr))
    first = anchor.find(A_R)
    props = first.find(qn("a:rPr")) if first is not None else anchor.find(qn("a:endParaRPr"))
    for el in _runs(text, props):
        p.append(el)
    end = anchor.find(qn("a:endParaRPr"))
    if end is not None:
        p.append(copy.deepcopy(end))
    return p


def fill(body: etree._Element, paragraphs: Sequence[Paragraph]) -> None:
    for p in body.findall(qn("a:p")):
        body.remove(p)
    for para in paragraphs:
        p = OxmlElement("a:p")
        if para.level:
            ppr = OxmlElement("a:pPr")
            ppr.set("lvl", str(para.level))
            p.append(ppr)
        for el in _runs(para.text, None):
            p.append(el)
        body.append(p)


def cell_paragraphs(tc: etree._Element) -> list[etree._Element]:
    body = tc.find(qn("a:txBody"))
    return body.findall(qn("a:p")) if body is not None else []


def cell_text(tc: etree._Element) -> str:
    return "\n".join(paragraph_text(p) for p in cell_paragraphs(tc))


def plan_cell(tc: etree._Element, new: str) -> list[tuple[etree._Element, Splice]]:
    paras = cell_paragraphs(tc)
    if not paras:
        raise SpliceError("the cell has no text body")
    lines = new.split("\n")
    return [(p, plan_splice(p, 0, paragraph_text(p), line)) for p, line in zip(paras[: len(lines)], lines[: len(paras)], strict=True)]


def write_cell(tc: etree._Element, new: str) -> None:
    paras = cell_paragraphs(tc)
    lines = new.split("\n")
    for p, s in plan_cell(tc, new):
        write_splice(p, s)
    last = paras[-1]
    for line in lines[len(paras) :]:
        p = new_paragraph(last, line)
        last.addnext(p)
        last = p
    for p in paras[len(lines) :]:
        p.getparent().remove(p)
