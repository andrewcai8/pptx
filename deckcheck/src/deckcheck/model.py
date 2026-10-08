from __future__ import annotations

import hashlib
import io
import re
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Literal

from lxml.etree import XMLSyntaxError
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.exc import PackageNotFoundError
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn
from pptx.shapes.picture import Picture

TITLE_TYPES = (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE)
OWN_MASTER_TYPES = (PP_PLACEHOLDER.DATE, PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.SLIDE_NUMBER, PP_PLACEHOLDER.HEADER)
MASTER_STYLES = {PP_PLACEHOLDER.TITLE: "p:titleStyle", PP_PLACEHOLDER.BODY: "p:bodyStyle"}
BULLET_TAGS = (qn("a:buNone"), qn("a:buChar"), qn("a:buAutoNum"))
INVISIBLE = "\u200b\u200c\u200d\ufeff"
LINE_BREAK = re.compile(r"[\v\n]")

Kind = Literal["title", "body", "text", "table", "chart", "picture", "other"]


class DeckError(Exception):
    pass


@dataclass(frozen=True)
class Run:
    text: str
    font: str
    size_pt: float | None


@dataclass(frozen=True)
class Paragraph:
    text: str
    runs: tuple[Run, ...]
    is_bullet: bool
    level: int

    @property
    def lines(self) -> list[str]:
        return [line.strip() for line in LINE_BREAK.split(self.text) if line.strip()]


@dataclass(frozen=True)
class Shape:
    name: str
    kind: Kind
    left: int
    top: int
    width: int
    height: int
    rotation: float
    paragraphs: tuple[Paragraph, ...]


@dataclass(frozen=True)
class Slide:
    index: int
    layout_name: str
    shapes: tuple[Shape, ...]

    @property
    def title_shape(self) -> Shape | None:
        return next((s for s in self.shapes if s.kind == "title"), None)

    @property
    def title(self) -> str:
        shape = self.title_shape
        return " ".join(p.text.strip() for p in shape.paragraphs) if shape else ""

    @property
    def headline(self) -> str:
        shape = self.title_shape
        return next((line for p in shape.paragraphs for line in p.lines), "") if shape else ""


@dataclass(frozen=True)
class Deck:
    path: str
    sha256: str
    slide_width: int
    slide_height: int
    slides: tuple[Slide, ...]


@dataclass(frozen=True)
class Violation:
    rule: str
    slide: int
    shape: str | None
    message: str
    evidence: str


@dataclass(frozen=True)
class ThemeFonts:
    major: str
    minor: str

    def resolve(self, typeface: str | None, is_title: bool) -> str:
        if typeface is None:
            return self.major if is_title else self.minor
        if typeface.startswith("+mj-"):
            return self.major
        if typeface.startswith("+mn-"):
            return self.minor
        return typeface


# x' = x * sx + dx, y' = y * sy + dy: maps group-child coordinates onto the slide.
@dataclass(frozen=True)
class Transform:
    sx: float = 1.0
    dx: float = 0.0
    sy: float = 1.0
    dy: float = 0.0


def load_deck(path: str | Path) -> Deck:
    try:
        data = Path(path).read_bytes()
        prs = Presentation(io.BytesIO(data))
    except (OSError, PackageNotFoundError, zipfile.BadZipFile, KeyError, ValueError, XMLSyntaxError) as e:
        raise DeckError(f"cannot read deck {path}: {e}") from e
    themes: dict[int, ThemeFonts] = {}
    slides = []
    for index, slide in enumerate(prs.slides, start=1):
        layout = slide.slide_layout
        theme = themes.setdefault(id(layout.slide_master.part), _theme_fonts(layout.slide_master))
        shapes = tuple(_shapes(slide.shapes, layout, theme, Transform()))
        slides.append(Slide(index, layout.name, shapes))
    return Deck(
        path=str(path),
        sha256=hashlib.sha256(data).hexdigest(),
        slide_width=int(prs.slide_width),
        slide_height=int(prs.slide_height),
        slides=tuple(slides),
    )


def _theme_fonts(master) -> ThemeFonts:
    theme = parse_xml(master.part.part_related_by(RT.THEME).blob)

    def latin(which: str) -> str:
        found = theme.xpath(f".//a:fontScheme/a:{which}/a:latin/@typeface")
        return str(found[0]) if found else ""

    return ThemeFonts(major=latin("majorFont"), minor=latin("minorFont"))


def _shapes(shapes, layout, theme: ThemeFonts, t: Transform):
    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _shapes(shape.shapes, layout, theme, _compose(t, shape))
            continue
        kind = _kind(shape)
        left, top = shape.left or 0, shape.top or 0
        width, height = shape.width or 0, shape.height or 0
        yield Shape(
            name=shape.name,
            kind=kind,
            left=round(left * t.sx + t.dx),
            top=round(top * t.sy + t.dy),
            width=round(width * t.sx),
            height=round(height * t.sy),
            rotation=shape.rotation,
            paragraphs=tuple(_paragraphs(shape, kind, theme, _bullet_styles(shape, layout))),
        )


def _compose(t: Transform, group) -> Transform:
    xfrm = group._element.grpSpPr.find(qn("a:xfrm"))
    if xfrm is None:
        return t
    off, ext = xfrm.find(qn("a:off")), xfrm.find(qn("a:ext"))
    ch_off, ch_ext = xfrm.find(qn("a:chOff")), xfrm.find(qn("a:chExt"))
    if off is None or ext is None or ch_off is None or ch_ext is None:
        return t
    sx = int(ext.get("cx")) / (int(ch_ext.get("cx")) or 1)
    sy = int(ext.get("cy")) / (int(ch_ext.get("cy")) or 1)
    dx = int(off.get("x")) - int(ch_off.get("x")) * sx
    dy = int(off.get("y")) - int(ch_off.get("y")) * sy
    return Transform(sx=sx * t.sx, dx=dx * t.sx + t.dx, sy=sy * t.sy, dy=dy * t.sy + t.dy)


def _kind(shape) -> Kind:
    if shape.has_chart:
        return "chart"
    if shape.has_table:
        return "table"
    if isinstance(shape, Picture):
        return "picture"
    if shape.is_placeholder:
        ph_type = shape.placeholder_format.type
        if ph_type in TITLE_TYPES:
            return "title"
        if ph_type in (PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT):
            return "body"
    if shape.has_text_frame:
        return "text"
    return "other"


# Reads go through oxml getters that return None for a missing element. python-pptx's cell.text_frame,
# run.font, and paragraph.level add the element they read, so reading through them edits the deck.
def _paragraphs(shape, kind: Kind, theme: ThemeFonts, styles: tuple):
    if kind == "table":
        bodies = [cell._tc.txBody for cell in shape.table.iter_cells()]
    elif shape.has_text_frame:
        bodies = [shape._element.txBody]
    else:
        bodies = []
    for body in bodies:
        for p in body.p_lst if body is not None else ():
            if not p.text.strip():
                continue
            level = p.pPr.lvl if p.pPr is not None else 0
            runs = tuple(_run(r, theme, kind == "title") for r in p.r_lst)
            yield Paragraph(text=p.text, runs=runs, is_bullet=_is_bullet(p, level, styles, kind), level=level)


def _run(r, theme: ThemeFonts, is_title: bool) -> Run:
    rPr = r.rPr
    latin = rPr.latin if rPr is not None else None
    sz = rPr.sz if rPr is not None else None
    return Run(
        text=r.text,
        font=theme.resolve(latin.typeface if latin is not None else None, is_title),
        size_pt=sz / 100 if sz is not None else None,
    )


# The list styles a paragraph inherits bullets from, nearest first: the shape's own, the matching
# layout placeholder's, the matching master placeholder's, then the master text style for that type.
def _bullet_styles(shape, layout) -> tuple:
    master = layout.slide_master
    styles = [_list_style(shape._element)]
    master_type = None
    if shape.is_placeholder:
        ph = shape._element
        base = next((p for p in layout.placeholders if p._element.ph_idx == ph.ph_idx), None) or next(
            (p for p in layout.placeholders if p._element.ph_type == ph.ph_type), None
        )
        master_type = _master_type(base._element.ph_type if base is not None else ph.ph_type)
        master_ph = next((p for p in master.placeholders if p._element.ph_type == master_type), None)
        styles += [_list_style(p._element) for p in (base, master_ph) if p is not None]
    tx_styles = master._element.find(qn("p:txStyles"))
    if tx_styles is not None:
        styles.append(tx_styles.find(qn(MASTER_STYLES.get(master_type, "p:otherStyle"))))
    return tuple(s for s in styles if s is not None)


def _master_type(ph_type) -> PP_PLACEHOLDER:
    if ph_type in TITLE_TYPES:
        return PP_PLACEHOLDER.TITLE
    return ph_type if ph_type in OWN_MASTER_TYPES else PP_PLACEHOLDER.BODY


def _list_style(element):
    body = element.find(qn("p:txBody"))
    return body.find(qn("a:lstStyle")) if body is not None else None


def _is_bullet(p, level: int, styles: tuple, kind: Kind) -> bool:
    tag = qn(f"a:lvl{level + 1}pPr")
    for pPr in (p.pPr, *(s.find(tag) for s in styles)):
        bullet = next(pPr.iterchildren(*BULLET_TAGS), None) if pPr is not None else None
        if bullet is not None:
            char = bullet.get("char", "")
            return bullet.tag == qn("a:buAutoNum") or any(not c.isspace() and c not in INVISIBLE for c in char)
    return kind == "body"
