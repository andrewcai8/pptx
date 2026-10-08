from pathlib import Path

import pytest
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches

from scenario import snapshot


def deck(path: Path, change=None) -> Path:
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Market outlook"
    slide.placeholders[1].text_frame.text = "Demand grows 4% a year"
    slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(5), Inches(3), Inches(1)).text_frame.text = "Callout"
    if change:
        change(slide)
    prs.save(path)
    return path


def seen(path: Path) -> tuple:
    snap = snapshot(path)
    return snap.looks, snap.shared


def paragraphs(slide):
    return [p for s in slide.shapes for p in s.text_frame.paragraphs]


def bullets_off(slide):
    for p in slide.placeholders[1].text_frame.paragraphs:
        p._p.get_or_add_pPr().insert_element_before(OxmlElement("a:buNone"), "a:tabLst", "a:defRPr", "a:extLst")


def box(slide):
    return slide.shapes[2]


READS = {
    "paragraph.level": lambda s: [p.level for p in paragraphs(s)],
    "paragraph.font": lambda s: [p.font.size for p in paragraphs(s)],
    "run.font": lambda s: [r.font.size for p in paragraphs(s) for r in p.runs],
    "shape.line.fill": lambda s: box(s).line.fill.type,
    "slide.notes_slide": lambda s: s.notes_slide.notes_text_frame.text,
}

EDITS = {
    "empty buNone": bullets_off,
    "empty noFill on a shape filled by its style": lambda s: box(s).fill.background(),
    "empty effectLst": lambda s: box(s)._element.spPr.append(OxmlElement("a:effectLst")),
    "normAutofit": lambda s: setattr(s.shapes.title.text_frame, "auto_size", MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE),
    "run.font.color": lambda s: [r.font.color for p in paragraphs(s) for r in p.runs],
    "shape.line.color": lambda s: box(s).line.color,
    "a notes page with text": lambda s: setattr(s.notes_slide.notes_text_frame, "text", "Say the number slowly"),
}


@pytest.mark.parametrize("read", READS.values(), ids=READS.keys())
def test_a_python_pptx_read_leaves_the_slide_unchanged(tmp_path, read):
    source = seen(deck(tmp_path / "source.pptx"))
    assert seen(deck(tmp_path / "read.pptx", read)) == source
    assert seen(deck(tmp_path / "edited.pptx", bullets_off)) != source


@pytest.mark.parametrize("edit", EDITS.values(), ids=EDITS.keys())
def test_an_edit_outside_the_read_artifacts_changes_the_slide(tmp_path, edit):
    assert seen(deck(tmp_path / "edited.pptx", edit)) != seen(deck(tmp_path / "source.pptx"))

