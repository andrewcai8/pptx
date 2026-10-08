import hashlib
import json
import shutil
import uuid
from decimal import Decimal
from pathlib import Path

import pytest
import yaml
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches

import score
from deckedit import DeckEdit, load_variants
from scenario import ROOT, BadScenario, Charts, open_scenario, snapshot


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


def scenario_at(d: Path, changes: list | None = None, make=deck) -> Path:
    d.mkdir(parents=True)
    source = make(d / "input.pptx")
    deck_ref = {"file": "input.pptx", "sha256": hashlib.sha256(source.read_bytes()).hexdigest()}
    (d / "expected.yaml").write_text(yaml.safe_dump({"deck": deck_ref, "changes": changes or []}))
    (d / "transcript.md").write_text("[00:00:05] Ana Ruiz (Principal, Kestrel Advisory): Nothing to change today.\n")
    return d


@pytest.fixture
def private_dir():
    top = ROOT / "private" / f"pytest-{uuid.uuid4().hex[:8]}"
    yield top
    shutil.rmtree(top, ignore_errors=True)


@pytest.mark.parametrize(("folder", "opens"), [("private", True), ("artifacts", False), ("evals", False)])
def test_a_private_deck_scenario_opens_only_under_the_repos_private_folder(folder, opens):
    top = ROOT / folder / f"pytest-{uuid.uuid4().hex[:8]}"
    try:
        d = scenario_at(top / "client-scenario")
        if opens:
            assert open_scenario(d).ref == "file:input.pptx"
        else:
            with pytest.raises(BadScenario, match="never enters git"):
                open_scenario(d)
    finally:
        shutil.rmtree(top, ignore_errors=True)


def edit_slide_1(may_change: list[str]) -> list[dict]:
    return [{"id": "c1", "kind": "edit-text", "intent": "Reword the callout.", "said": ["00:00:05"], "slides": {1: {"may_change": may_change}}}]


def test_may_change_names_a_text_shape_on_the_edited_slide(private_dir):
    sc = open_scenario(scenario_at(private_dir / "callout", edit_slide_1(["Rectangle 3"])))
    assert sc.changes[0].slides[1].may_change == ("Rectangle 3",)
    with pytest.raises(BadScenario, match="slide 1 has no text shape named 'Footnote'"):
        open_scenario(scenario_at(private_dir / "footnote", edit_slide_1(["Footnote"])))


def chart_deck(path: Path, plot: str = "barChart") -> Path:
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    data = CategoryChartData()
    data.categories = ["2022", "2027"]
    data.add_series("Market", (410, 2000))
    chart = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(2), Inches(6), Inches(4), data).chart
    chart._chartSpace.plotArea.find(qn("c:barChart")).tag = qn(f"c:{plot}")
    prs.save(path)
    return path


@pytest.mark.parametrize("plot", ["bar3DChart", "line3DChart", "pie3DChart", "stockChart", "surfaceChart", "surface3DChart", "ofPieChart"])
def test_a_chart_python_pptx_cannot_read_is_reported_by_type(tmp_path, plot):
    assert snapshot(chart_deck(tmp_path / "unreadable.pptx", plot)).charts == (Charts((), (plot,)),)
    assert snapshot(chart_deck(tmp_path / "bar.pptx")).charts == (Charts((Decimal("410.0"), Decimal("2000.0")), ()),)


@pytest.mark.parametrize(
    ("variant", "code", "verdict", "failures"),
    [
        ("chart_drawn_in_3d", score.BAD, "UNREADABLE", [("unreadable", 10, "slide 10 chart: bar3DChart cannot be read; check by hand")]),
        ("untouched_chart_drawn_in_3d", score.FAIL, "FAIL", [("scope", 12, "slide 12 changed a chart, image, media, or notes part, but no change asks for it")]),
    ],
)
def test_a_3d_chart_on_a_copy_of_the_solar_deck_is_scored_without_a_crash(tmp_path, variant, code, verdict, failures):
    sc = open_scenario("solar-market-refresh")
    d = DeckEdit(sc)
    next(v for v in load_variants(sc.dir / "build.py") if v.name == variant).apply(d)
    d.save(tmp_path / "output.pptx")
    assert score.main(["solar-market-refresh", str(tmp_path / "output.pptx")]) == code
    report = json.loads((tmp_path / "score.json").read_text())
    assert report["verdict"] == verdict
    assert [(f["code"], f["slide"], f["message"]) for f in report["failures"]] == failures


def test_a_chart_fact_cannot_target_a_source_chart_that_cannot_be_read(private_dir):
    change = [{"id": "c1", "kind": "update-number", "intent": "Redraw the bar.", "said": ["00:00:05"], "slides": {1: {"forbid": [{"chart": 410}]}}}]
    assert open_scenario(scenario_at(private_dir / "bar", change, chart_deck)).source.charts[0].values == (Decimal("410.0"), Decimal("2000.0"))
    with pytest.raises(BadScenario, match="c1 slide 1: a chart fact cannot target a slide whose bar3DChart chart cannot be read"):
        open_scenario(scenario_at(private_dir / "bar3d", change, lambda path: chart_deck(path, "bar3DChart")))
