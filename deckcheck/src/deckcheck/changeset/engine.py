from __future__ import annotations

import hashlib
import json
import zipfile
from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from pathlib import Path

from lxml import etree
from pptx.oxml.ns import qn
from pydantic import ValidationError

from deckcheck.changeset import chart, slides, text
from deckcheck.changeset.model import (
    OP_KINDS,
    AddSlide,
    Box,
    Change,
    ChangeSet,
    DeleteSlide,
    Edited,
    Executed,
    FillPlaceholder,
    InsertParagraph,
    MoveSlide,
    Op,
    ReplaceText,
    Review,
    ReviewChange,
    ReviewShape,
    ReviewSlide,
    SetCell,
    SetChartValue,
)
from deckcheck.model import DeckError, open_presentation, read_deck
from deckcheck.package import Package

DERIVED = {
    "before": "the engine reads this from the source deck; remove it",
    "after": "the engine derives this from the op; remove it",
    "structural": "the engine derives this from the op kind; remove it",
    "depends_on": "the engine derives this from the op; remove it",
}
DECISION_FORMAT = 'must be "pending", "keep_new", "keep_old" or {"edited": "<text>"}'


@dataclass(frozen=True)
class Problem:
    where: str
    message: str


class Invalid(Exception):
    def __init__(self, problems: Sequence[Problem]) -> None:
        super().__init__(f"{len(problems)} problems")
        self.problems = tuple(problems)


class Undecided(Exception):
    def __init__(self, pending: Sequence[str]) -> None:
        super().__init__(f"{len(pending)} pending")
        self.pending = tuple(pending)


class Miss(Exception):
    def __init__(self, field: str, message: str) -> None:
        super().__init__(message)
        self.field = field
        self.message = message


@dataclass(frozen=True)
class TextAt:
    slide: slides.SourceSlide
    shape: etree._Element
    p: etree._Element
    paragraph: int
    at: int
    before: str
    splice: text.Splice


@dataclass(frozen=True)
class ParagraphAt:
    slide: slides.SourceSlide
    shape: etree._Element
    anchor: etree._Element


@dataclass(frozen=True)
class CellAt:
    slide: slides.SourceSlide
    shape: etree._Element
    tc: etree._Element


@dataclass(frozen=True)
class PointAt:
    slide: slides.SourceSlide
    shape: etree._Element
    point: chart.Point


@dataclass(frozen=True)
class NewSlide:
    layout: slides.Layout
    id: int
    part: str
    after: slides.SourceSlide | None


@dataclass(frozen=True)
class PlaceholderAt:
    add: NewSlide
    placeholder: slides.Placeholder


@dataclass(frozen=True)
class SlideAt:
    slide: slides.SourceSlide


@dataclass(frozen=True)
class MoveTo:
    slide: slides.SourceSlide
    after: slides.SourceSlide | None


Target = TextAt | ParagraphAt | CellAt | PointAt | NewSlide | PlaceholderAt | SlideAt | MoveTo


@dataclass(frozen=True)
class Item:
    """One change as checked against the source deck. Everything here is read from the source, not the maker."""

    change: Change
    slide: int | str
    source_index: int | None
    shape: tuple[int, str] | None
    before: str | None
    after: str | None
    span: tuple[int, int] | None
    depends_on: str | None
    notes: tuple[str, ...]


@dataclass(frozen=True)
class Checked:
    """A ChangeSet proved against its source deck: every id resolves, every quote matches, nothing conflicts,
    and every decision is admissible, so execute cannot fail and apply fails only on a pending decision."""

    path: Path
    changeset: ChangeSet
    source: bytes
    items: tuple[Item, ...]


@dataclass(frozen=True)
class Applied:
    data: bytes
    kept_new: tuple[str, ...]
    edited: tuple[str, ...]
    kept_old: tuple[str, ...]
    dropped: tuple[tuple[str, str], ...]


@dataclass(frozen=True)
class Context:
    """What locating one op needs beyond the deck: every change's op and each add_slide's ordinal among all
    add_slide changes, so an added slide has the same id and part whichever adds a reviewer keeps."""

    ops: Mapping[str, Op]
    ordinals: Mapping[str, int]

    @classmethod
    def of(cls, cs: ChangeSet) -> Context:
        adds = [c.id for c in cs.changes if isinstance(c.op, AddSlide)]
        return cls({c.id: c.op for c in cs.changes}, {cid: i for i, cid in enumerate(adds, start=1)})


def load(path: Path) -> Checked:
    try:
        raw_text = path.read_text(encoding="utf-8-sig")
    except (OSError, UnicodeDecodeError) as e:
        raise DeckError(f"cannot read ChangeSet {path}: {e}") from e
    cs = parse(raw_text)
    source = Path(cs.source.path)
    try:
        data = source.read_bytes()
    except OSError as e:
        raise Invalid([Problem("source.path", f"cannot read {source}: {e.strerror}")]) from e
    sha = hashlib.sha256(data).hexdigest()
    if sha != cs.source.sha256:
        raise Invalid([Problem("source.sha256", f"{source} has sha256 {sha}; the deck changed since the maker read it")])
    return check(cs, data, path)


def parse(raw_text: str) -> ChangeSet:
    try:
        raw = json.loads(raw_text)
    except json.JSONDecodeError as e:
        raise Invalid([Problem("changeset", f"not JSON: {e}")]) from e
    try:
        return ChangeSet.model_validate_json(raw_text)
    except ValidationError as e:
        problems: dict[str, Problem] = {}
        for err in e.errors():
            where = _where(err["loc"], raw)
            problems.setdefault(where, Problem(where, _message(err)))
        raise Invalid(list(problems.values())) from e


def _where(loc: tuple, raw: object) -> str:
    if len(loc) < 2 or loc[0] != "changes" or not isinstance(loc[1], int):
        return ".".join(map(str, loc)) or "changeset"
    change = raw["changes"][loc[1]] if isinstance(raw, dict) and isinstance(raw.get("changes"), list) else None
    cid = change.get("id") if isinstance(change, dict) else None
    name = cid if isinstance(cid, str) and cid else f"changes[{loc[1]}]"
    rest = list(loc[2:])
    if rest[:1] == ["op"] and len(rest) > 1 and rest[1] in OP_KINDS:
        del rest[1]
    if rest[:1] == ["decision"]:
        rest = ["decision"]
    return " ".join([name, ".".join(map(str, rest))]) if rest else name


def _message(err: Mapping) -> str:
    field = err["loc"][-1] if err["loc"] else ""
    if err["type"] == "extra_forbidden":
        return DERIVED.get(str(field), "unknown field; remove it")
    if "decision" in err["loc"]:
        return DECISION_FORMAT
    if err["type"] == "missing":
        return "required"
    return err["msg"]


def check(cs: ChangeSet, data: bytes, path: Path = Path("changeset.json")) -> Checked:
    try:
        pkg = Package(data)
        deck = slides.read_deck(pkg)
    except (zipfile.BadZipFile, KeyError, StopIteration, etree.XMLSyntaxError) as e:
        raise DeckError(f"cannot read deck {cs.source.path}: {e}") from e
    ctx = Context.of(cs)
    problems: list[Problem] = []
    targets: dict[str, Target] = {}
    for change in cs.changes:
        try:
            targets[change.id] = locate(pkg, deck, ctx, change.id, change.op)
        except Miss as m:
            problems.append(Problem(f"{change.id} {m.field}", m.message))
            continue
        problems += _decision_problems(pkg, deck, ctx, change)
    problems += _conflicts(cs, deck, targets)
    if problems:
        raise Invalid(problems)
    return Checked(path, cs, data, tuple(_item(c, targets[c.id], deck, pkg) for c in cs.changes))


def _decision_problems(pkg: Package, deck: slides.Deck, ctx: Context, change: Change) -> list[Problem]:
    if not isinstance(change.decision, Edited):
        return []
    try:
        op = change.op.edit(change.decision.edited)
        locate(pkg, deck, ctx, change.id, op)
    except ValueError as e:
        return [Problem(f"{change.id} decision", str(e))]
    except Miss as m:
        return [Problem(f"{change.id} decision", f"the edited text cannot be written: {m.message}")]
    return []


def locate(pkg: Package, deck: slides.Deck, ctx: Context, cid: str, op: Op) -> Target:
    """Find what `op` writes on a pristine package, or raise Miss naming the field at fault. check and _build
    both call this, so a ChangeSet that checks cannot fail to build."""
    match op:
        case ReplaceText():
            return _locate_text(pkg, deck, op)
        case InsertParagraph():
            slide, shape, paras = _paragraphs(pkg, deck, op.slide, op.shape)
            if op.after >= len(paras):
                raise Miss("op.after", f"shape {op.shape} {slides.shape_name(shape)!r} has {_count(len(paras))}")
            _no_newline(op.text, "op.text")
            return ParagraphAt(slide, shape, paras[op.after])
        case SetCell():
            return _locate_cell(pkg, deck, op)
        case SetChartValue():
            slide = _slide(deck, op.slide, "op.slide")
            shape = _shape(pkg, slide, op.shape)
            try:
                point = chart.locate_point(pkg, slide.part, shape, op.series, op.point)
            except chart.ChartError as e:
                raise Miss("op", f"shape {op.shape} {slides.shape_name(shape)!r}: {e}") from e
            if point.value != op.old:
                raise Miss("op.old", f"series {op.series} point {op.point} caches {chart.number_text(point.value)}, not {chart.number_text(op.old)}")
            return PointAt(slide, shape, point)
        case AddSlide():
            return _locate_add(deck, ctx, cid, op)
        case FillPlaceholder():
            return _locate_fill(pkg, deck, ctx, op)
        case DeleteSlide():
            return SlideAt(_slide(deck, op.slide, "op.slide"))
        case MoveSlide():
            after = _slide(deck, op.after, "op.after") if op.after is not None else None
            return MoveTo(_slide(deck, op.slide, "op.slide"), after)
    raise AssertionError(op)


def _locate_text(pkg: Package, deck: slides.Deck, op: ReplaceText) -> TextAt:
    _no_newline(op.old, "op.old")
    _no_newline(op.new, "op.new")
    slide, shape, paras = _paragraphs(pkg, deck, op.slide, op.shape)
    name = f"shape {op.shape} {slides.shape_name(shape)!r}"
    if op.paragraph is not None and op.paragraph >= len(paras):
        raise Miss("op.paragraph", f"{name} has {_count(len(paras))}")
    texts = [text.paragraph_text(p) for p in paras]
    hits = [(i, at) for i, t in enumerate(texts) for at in text.occurrences(t, op.old)]
    mine = [(i, at) for i, at in hits if op.paragraph in (None, i)]
    if not mine:
        where = f" paragraph {op.paragraph}" if op.paragraph is not None else ""
        elsewhere = f"; it occurs in {_paragraph_list({i for i, _ in hits})}" if hits else ""
        raise Miss("op.old", f"{op.old!r} is not in {name}{where}{elsewhere}; its text is {_clip(chr(10).join(texts))!r}")
    if len(mine) > 1:
        raise Miss(
            "op.old",
            f"{op.old!r} occurs {len(mine)} times in {name}, in {_paragraph_list({i for i, _ in mine})}; "
            "quote more of the text or name the paragraph",
        )
    i, at = mine[0]
    try:
        splice = text.plan_splice(paras[i], at, op.old, op.new)
    except text.SpliceError as e:
        raise Miss("op.new", str(e)) from e
    return TextAt(slide, shape, paras[i], i, at, texts[i], splice)


def _locate_cell(pkg: Package, deck: slides.Deck, op: SetCell) -> CellAt:
    slide = _slide(deck, op.slide, "op.slide")
    shape = _shape(pkg, slide, op.shape)
    tbl = shape.find(f".//{qn('a:tbl')}") if shape.tag == qn("p:graphicFrame") else None
    if tbl is None:
        raise Miss("op.shape", f"shape {op.shape} {slides.shape_name(shape)!r} is not a table")
    rows = tbl.findall(qn("a:tr"))
    if op.row >= len(rows):
        raise Miss("op.row", f"the table has {len(rows)} rows, numbered 0 to {len(rows) - 1}")
    cells = rows[op.row].findall(qn("a:tc"))
    if op.col >= len(cells):
        raise Miss("op.col", f"row {op.row} has {len(cells)} cells, numbered 0 to {len(cells) - 1}")
    tc = cells[op.col]
    if tc.get("hMerge") in ("1", "true") or tc.get("vMerge") in ("1", "true"):
        raise Miss("op.col", f"row {op.row} col {op.col} is merged into a neighbouring cell; set that cell")
    source = text.cell_text(tc)
    if source != op.old:
        raise Miss("op.old", f"row {op.row} col {op.col} reads {source!r}, not {op.old!r}")
    try:
        text.plan_cell(tc, op.new)
    except text.SpliceError as e:
        raise Miss("op.new", str(e)) from e
    return CellAt(slide, shape, tc)


def _locate_add(deck: slides.Deck, ctx: Context, cid: str, op: AddSlide) -> NewSlide:
    found = [layout for layout in deck.layouts if layout.name == op.layout]
    if not found:
        names = sorted({layout.name for layout in deck.layouts})
        raise Miss("op.layout", f"no layout {op.layout!r} on the masters the deck's slides use; layouts: {_list(map(repr, names))}")
    if len(found) > 1:
        raise Miss("op.layout", f"{len(found)} masters have a layout named {op.layout!r}, so the name does not say which")
    after = _slide(deck, op.after, "op.after") if op.after is not None else None
    ordinal = ctx.ordinals[cid]
    slide_id = max((s.id for s in deck.slides), default=255) + ordinal
    if slide_id > slides.MAX_SLIDE_ID:
        raise Miss("op", "no slide id is free above the deck's highest")
    return NewSlide(found[0], slide_id, f"ppt/slides/slide{deck.last_part + ordinal}.xml", after)


def _locate_fill(pkg: Package, deck: slides.Deck, ctx: Context, op: FillPlaceholder) -> PlaceholderAt:
    add = ctx.ops.get(op.slide)
    if add is None:
        raise Miss("op.slide", f"no change {op.slide!r}; fill_placeholder.slide names an add_slide change")
    if not isinstance(add, AddSlide):
        raise Miss("op.slide", f"{op.slide} is a {add.kind} change, not an add_slide")
    try:
        new = _locate_add(deck, ctx, op.slide, add)
    except Miss as m:
        raise Miss("op.slide", f"{op.slide} cannot add its slide: {m.message}") from m
    ph = next((p for p in new.layout.placeholders if p.id == op.shape), None)
    if ph is None:
        listed = _list(f"{p.id} {p.name!r} ({p.type})" for p in new.layout.placeholders)
        raise Miss("op.shape", f"layout {new.layout.name!r} has no placeholder {op.shape}; its placeholders: {listed}")
    if not ph.holds_text:
        raise Miss("op.shape", f"placeholder {op.shape} {ph.name!r} is a {ph.type} placeholder and holds no text")
    for i, para in enumerate(op.paragraphs):
        _no_newline(para.text, f"op.paragraphs.{i}.text")
    return PlaceholderAt(new, ph)


def _slide(deck: slides.Deck, slide_id: int, field: str) -> slides.SourceSlide:
    slide = deck.slide(slide_id)
    if slide is None:
        raise Miss(field, _no_slide(deck, slide_id))
    return slide


def _no_slide(deck: slides.Deck, slide_id: int) -> str:
    return f"no slide {slide_id}; slide ids: {_list(f'{s.id} (slide {s.index})' for s in deck.slides)}"


def _shape(pkg: Package, slide: slides.SourceSlide, shape_id: int) -> etree._Element:
    found = [s for s in slides.shapes(slides.shape_tree(pkg.xml(slide.part))) if slides.shape_id(s) == shape_id]
    if len(found) == 1:
        return found[0]
    if found:
        raise Miss("op.shape", f"{slide.name} has {len(found)} shapes with id {shape_id}, so the id does not say which")
    listed = []
    for s in slides.shapes(slides.shape_tree(pkg.xml(slide.part))):
        body = s.find(qn("p:txBody"))
        words = " ".join(text.paragraph_text(p) for p in body.iterfind(qn("a:p"))).strip() if body is not None else ""
        if words:
            listed.append(f"{slides.shape_id(s)} {slides.shape_name(s)!r} ({_clip(words, 30)!r})")
    raise Miss("op.shape", f"{slide.name} has no shape {shape_id}; shapes with text: {_list(listed)}")


def _paragraphs(pkg: Package, deck: slides.Deck, slide_id: int, shape_id: int):
    slide = _slide(deck, slide_id, "op.slide")
    shape = _shape(pkg, slide, shape_id)
    body = shape.find(qn("p:txBody"))
    if body is None:
        hint = "; use set_cell" if shape.find(f".//{qn('a:tbl')}") is not None else ""
        hint = hint or ("; use set_chart_value" if shape.find(f".//{{{chart.C_NS}}}chart") is not None else "")
        raise Miss("op.shape", f"shape {shape_id} {slides.shape_name(shape)!r} holds no text{hint}")
    return slide, shape, body.findall(qn("a:p"))


def _no_newline(value: str, field: str) -> None:
    if "\n" in value:
        raise Miss(field, '"\\n" would start a new paragraph; use insert_paragraph, or "\\v" for a line break')


def _paragraph_list(indices: set[int]) -> str:
    return f"{'paragraph' if len(indices) == 1 else 'paragraphs'} {_list(sorted(indices))}"


def _count(n: int) -> str:
    return f"{n} paragraphs, numbered 0 to {n - 1}" if n else "no paragraphs"


def _list(items) -> str:
    return ", ".join(str(i) for i in items)


def _clip(s: str, n: int = 120) -> str:
    return s if len(s) <= n else s[: n - 1] + "…"


def _conflicts(cs: ChangeSet, deck: slides.Deck, targets: Mapping[str, Target]) -> list[Problem]:
    """Judged over every change whatever its decision, so any subset a reviewer keeps is conflict-free."""
    problems: list[Problem] = []
    for kind, items in (("changes", cs.changes), ("asks", cs.asks), ("flags", cs.flags), ("held", cs.held)):
        first: dict[str, int] = {}
        for i, item in enumerate(items):
            if item.id in first:
                problems.append(Problem(f"{item.id} id", f"{kind}[{first[item.id]}] and {kind}[{i}] share this id"))
            first.setdefault(item.id, i)
    for item in (*cs.flags, *cs.held):
        for slide_id in item.slides:
            if deck.slide(slide_id) is None:
                problems.append(Problem(f"{item.id} slides", _no_slide(deck, slide_id)))
    asks = [a.id for a in cs.asks]
    for c in cs.changes:
        if c.ask_id not in asks:
            problems.append(Problem(f"{c.id} ask_id", f"no ask {c.ask_id!r}; asks: {_list(asks) or 'none'}"))

    def name(slide_id: int) -> str:
        slide = deck.slide(slide_id)
        return slide.name if slide else f"slide {slide_id}"

    deleted: dict[int, str] = {}
    for c in cs.changes:
        if isinstance(c.op, DeleteSlide):
            if c.op.slide in deleted:
                problems.append(Problem(f"{c.id} op.slide", f"{name(c.op.slide)} is already deleted by {deleted[c.op.slide]}"))
            deleted.setdefault(c.op.slide, c.id)
    moved: dict[int, str] = {}
    claims: dict[tuple, str] = {}
    quotes: dict[tuple, list[tuple[int, int, str, str]]] = {}
    for c in cs.changes:
        op = c.op
        if isinstance(op, MoveSlide):
            if op.slide in moved:
                problems.append(Problem(f"{c.id} op.slide", f"{name(op.slide)} is already moved by {moved[op.slide]}"))
            if op.after == op.slide:
                problems.append(Problem(f"{c.id} op.after", "a slide cannot follow itself"))
            moved.setdefault(op.slide, c.id)
        if isinstance(op, ReplaceText | InsertParagraph | SetCell | SetChartValue | MoveSlide) and op.slide in deleted:
            problems.append(Problem(f"{c.id} op.slide", f"{name(op.slide)} is deleted by {deleted[op.slide]}"))
        if isinstance(op, AddSlide | MoveSlide) and op.after in deleted:
            problems.append(
                Problem(f"{c.id} op.after", f"{name(op.after)} is deleted by {deleted[op.after]}; nothing can follow it")
            )
        key = _claim(op)
        if key in claims:
            problems.append(Problem(f"{c.id} op", f"{_claim_name(key)} is already changed by {claims[key]}"))
        elif key is not None:
            claims[key] = c.id
        target = targets.get(c.id)
        if isinstance(target, TextAt):
            spans = quotes.setdefault((op.slide, op.shape, target.paragraph), [])
            end = target.at + len(op.old)
            other = next((s for s in spans if target.at < s[1] and s[0] < end), None)
            if other:
                problems.append(
                    Problem(f"{c.id} op.old", f"{op.old!r} overlaps {other[2]}'s quote {other[3]!r} in paragraph {target.paragraph}")
                )
            spans.append((target.at, end, c.id, op.old))
    problems += _cycles(cs, name)
    return problems


def _cycles(cs: ChangeSet, name) -> list[Problem]:
    after = {c.op.slide: (c.op.after, c.id) for c in cs.changes if isinstance(c.op, MoveSlide)}
    problems = []
    for start, (_, cid) in after.items():
        seen, at = {start}, after[start][0]
        while at in after and at not in seen:
            seen.add(at)
            at = after[at][0]
        if at == start and after[start][0] != start:
            problems.append(Problem(f"{cid} op.after", f"{name(start)} is placed after a slide that is placed after it"))
    return problems


def _claim(op: Op) -> tuple | None:
    match op:
        case SetCell():
            return ("cell", op.slide, op.shape, op.row, op.col)
        case SetChartValue():
            return ("point", op.slide, op.shape, op.series, op.point)
        case FillPlaceholder():
            return ("placeholder", op.slide, op.shape)
    return None


def _claim_name(key: tuple) -> str:
    match key:
        case ("cell", _, shape, row, col):
            return f"shape {shape} row {row} col {col}"
        case ("point", _, shape, series, point):
            return f"shape {shape} series {series} point {point}"
        case ("placeholder", add, shape):
            return f"placeholder {shape} on {add}'s slide"
    raise AssertionError(key)


def _item(c: Change, target: Target, deck: slides.Deck, pkg: Package) -> Item:
    op = c.op
    match target:
        case TextAt():
            after = target.before[: target.at] + op.new + target.before[target.at + len(op.old) :]
            notes = ()
            if any(x.el.tag == text.A_FLD for x in text.atoms(target.p) if x.lo < target.splice.end and x.hi > target.splice.start):
                notes = ("a think-cell label: its field format is rewritten too, so a refresh keeps the new text",)
            return _at(c, target.slide, target.shape, target.before, after, (target.at, target.at + len(op.old)), notes)
        case ParagraphAt():
            return _at(c, target.slide, target.shape, None, op.text)
        case CellAt():
            return _at(c, target.slide, target.shape, op.old, op.new)
        case PointAt():
            book = target.point.workbook
            notes = (f"workbook cell {target.point.cell.sheet}!{target.point.cell.ref} in {book}",)
            if book.endswith(".xlsb"):
                notes += ("the .xlsb workbook is rewritten as .xlsx, values only",)
            return _at(c, target.slide, target.shape, chart.number_text(op.old), chart.number_text(op.new), None, notes)
        case PlaceholderAt():
            after = "\n".join(p.text for p in op.paragraphs)
            return Item(c, op.slide, None, (target.placeholder.id, target.placeholder.name), None, after, None, op.slide, ())
        case SlideAt():
            return Item(c, target.slide.id, target.slide.index, None, slides.title(pkg.xml(target.slide.part)), None, None, None, ())
        case MoveTo():
            before = f"after {deck.slides[target.slide.index - 2].name}" if target.slide.index > 1 else "first"
            after = f"after {target.after.name}" if target.after else "first"
            return Item(c, target.slide.id, target.slide.index, None, before, after, None, None, ())
        case NewSlide():
            where = f"after {target.after.name}" if target.after else "first"
            return Item(c, c.id, None, None, None, f"new slide on layout {target.layout.name!r}, {where}", None, None, ())
    raise AssertionError(target)


def _at(c, slide, shape, before, after, span=None, notes=()) -> Item:
    return Item(c, slide.id, slide.index, (slides.shape_id(shape), slides.shape_name(shape)), before, after, span, None, notes)


def execute(checked: Checked) -> bytes:
    """The maker-output deck: every change as written. Equal to apply with every decision keep_new."""
    return _build(checked, [(c.id, c.op) for c in checked.changeset.changes])


def apply(checked: Checked) -> Applied:
    """The final deck, from a fresh copy of the source: keep_new and edited changes replayed, keep_old ones not."""
    changes = checked.changeset.changes
    pending = [c.id for c in changes if c.decision == "pending"]
    if pending:
        raise Undecided(pending)
    kept_adds = {c.id for c in changes if isinstance(c.op, AddSlide) and c.decision != "keep_old"}
    ops: list[tuple[str, Op]] = []
    kept_new, edited, kept_old, dropped = [], [], [], []
    for c in changes:
        if c.decision == "keep_old":
            kept_old.append(c.id)
            continue
        if isinstance(c.op, FillPlaceholder) and c.op.slide not in kept_adds:
            dropped.append((c.id, c.op.slide))
            continue
        if isinstance(c.decision, Edited):
            ops.append((c.id, c.op.edit(c.decision.edited)))
            edited.append(c.id)
        else:
            ops.append((c.id, c.op))
            kept_new.append(c.id)
    return Applied(_build(checked, ops), tuple(kept_new), tuple(edited), tuple(kept_old), tuple(dropped))


def _build(checked: Checked, ops: Sequence[tuple[str, Op]]) -> bytes:
    """The one write path: a pure function of the source bytes and the ops. Every target is located on the
    pristine trees before anything is written, so no write shifts another change's address."""
    pkg = Package(checked.source)
    deck = slides.read_deck(pkg)
    ctx = Context.of(checked.changeset)
    located = [(cid, op, locate(pkg, deck, ctx, cid, op)) for cid, op in ops]
    spans = [t for _, _, t in located if isinstance(t, TextAt)]
    # Quotes in one paragraph never overlap, so their starts order them. Splice offsets do not: two abutting
    # quotes can both trim to an insertion at their shared edge, and the later quote's must be written first.
    for t in sorted(spans, key=lambda t: t.at, reverse=True):
        text.write_splice(t.p, t.splice)
    tails: dict[etree._Element, etree._Element] = {}
    rids: dict[int, str] = {s.id: s.rid for s in deck.slides}
    placements: list[tuple[int, int | None]] = []
    deleted: set[int] = set()
    for _, op, t in located:
        match t:
            case ParagraphAt():
                p = text.new_paragraph(t.anchor, op.text)
                tails.get(t.anchor, t.anchor).addnext(p)
                tails[t.anchor] = p
            case CellAt():
                text.write_cell(t.tc, op.new)
            case PointAt():
                chart.set_point(pkg, t.point, op.new)
            case NewSlide():
                rids[t.id] = slides.new_slide(pkg, deck, t.layout, t.part)
                placements.append((t.id, t.after.id if t.after else None))
            case SlideAt():
                deleted.add(t.slide.id)
            case MoveTo():
                placements.append((t.slide.id, t.after.id if t.after else None))
    for _, op, t in located:
        if isinstance(t, PlaceholderAt):
            sp = next(s for s in slides.shapes(slides.shape_tree(pkg.xml(t.add.part))) if slides.shape_id(s) == t.placeholder.id)
            text.fill(sp.find(qn("p:txBody")), op.paragraphs)
    order = slides.final_order([s.id for s in deck.slides], deleted, placements)
    slides.write_order(pkg, deck, [(sid, rids[sid]) for sid in order], dict(placements))
    return pkg.to_bytes()


def review(checked: Checked, executed_path: Path, executed: bytes) -> Review:
    cs = checked.changeset
    pkg = Package(checked.source)
    deck = slides.read_deck(pkg)
    adds = {i.change.id: i for i in checked.items if isinstance(i.change.op, AddSlide)}
    placements = [
        (i.change.id if isinstance(i.change.op, AddSlide) else i.change.op.slide, i.change.op.after)
        for i in checked.items
        if isinstance(i.change.op, AddSlide | MoveSlide)
    ]
    deleted = {i.change.op.slide for i in checked.items if isinstance(i.change.op, DeleteSlide)}
    order = slides.final_order([s.id for s in deck.slides], deleted, placements)
    executed_index = {key: n for n, key in enumerate(order, start=1)}
    ctx = Context.of(cs)
    titles = {
        i.change.op.slide: (i.after or "").split("\n")[0]
        for i in checked.items
        if isinstance(i.change.op, FillPlaceholder)
        and _locate_fill(pkg, deck, ctx, i.change.op).placeholder.type in slides.TITLE_TYPES
    }
    boxes = _boxes(checked.source, cs.source.path)
    slide_rows = [
        ReviewSlide(key=s.id, source_index=s.index, executed_index=executed_index.get(s.id), title=slides.title(pkg.xml(s.part)))
        for s in deck.slides
    ] + [ReviewSlide(key=cid, source_index=None, executed_index=executed_index[cid], title=titles.get(cid, "")) for cid in adds]
    changes = [
        ReviewChange(
            id=i.change.id,
            ask_id=i.change.ask_id,
            kind=i.change.op.kind,
            structural=i.change.op.structural,
            slide=i.slide,
            source_index=i.source_index,
            executed_index=executed_index.get(i.slide),
            shape=_review_shape(i, boxes),
            before=i.before,
            after=i.after,
            span=i.span,
            depends_on=i.depends_on,
            admits=i.change.op.admits,
            decision=i.change.decision,
            rationale=i.change.rationale,
            refs=i.change.refs,
            notes=i.notes,
        )
        for i in checked.items
    ]
    return Review(
        source=cs.source,
        executed=Executed(path=str(executed_path), sha256=hashlib.sha256(executed).hexdigest()),
        meeting=cs.meeting,
        asks=cs.asks,
        flags=cs.flags,
        held=cs.held,
        slides=tuple(slide_rows),
        changes=tuple(changes),
    )


def _boxes(data: bytes, path: str) -> dict[tuple[int, int], Box]:
    deck = read_deck(open_presentation(data, path), path, "")
    return {
        (slide.index, shape.xml.shape_id): Box(x=shape.left, y=shape.top, w=shape.width, h=shape.height)
        for slide in deck.slides
        for shape in slide.shapes
    }


def _review_shape(item: Item, boxes: Mapping[tuple[int, int], Box]) -> ReviewShape | None:
    if item.shape is None:
        return None
    shape_id, name = item.shape
    box = boxes.get((item.source_index, shape_id)) if item.source_index is not None else None
    return ReviewShape(id=shape_id, name=name, box=box)
