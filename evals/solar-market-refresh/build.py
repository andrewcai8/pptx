from pptx.oxml.ns import qn

from deckedit import DeckEdit, variant


def _cagr_label(d: DeckEdit, old: str, new: str) -> None:
    fields = [f for f in d.slide(10)._element.iter(qn("a:fld")) if f.findtext(qn("a:t")) == f"+{old}%"]
    if len(fields) != 1:
        raise ValueError(f"slide 10: expected one +{old}% think-cell label, found {len(fields)}")
    field = fields[0]
    field.find(qn("a:t")).text = f"+{new}%"
    field.set("type", field.get("type").replace(f"'{old[1]}'", f"'{new[1]}'"))


def _chart_2022_bar(d: DeckEdit, old: str, new: str) -> None:
    chart = next(s.chart for s in d.slide(10).shapes if s.has_chart)
    cells = [v for v in chart._chartSpace.iter(qn("c:v")) if v.text in (old, f"{old}.0")]
    if len(cells) != 1:
        raise ValueError(f"slide 10 chart: expected one {old} data point, found {len(cells)}")
    cells[0].text = new


# python-pptx reads no 3D plot, so a chart drawn in 3D is one the scorer cannot read.
def _draw_in_3d(d: DeckEdit, slide: int) -> None:
    chart = next(s.chart for s in d.slide(slide).shapes if s.has_chart)
    plot = chart._chartSpace.plotArea.find(qn("c:barChart"))
    for el in plot.findall(qn("c:overlap")) + plot.findall(qn("c:serLines")):
        plot.remove(el)
    plot.tag = qn("c:bar3DChart")


def _table_cell(d: DeckEdit, old: str, new: str) -> None:
    table = next(s.table for s in d.slide(10).shapes if s.has_table)
    runs = [r for row in table.rows for c in row.cells for p in c.text_frame.paragraphs for r in p.runs if r.text == old]
    if len(runs) != 1:
        raise ValueError(f"slide 10 table: expected one {old} cell, found {len(runs)}")
    runs[0].text = new


def _refresh(d: DeckEdit, *, size: str = "410", contents: bool = True) -> None:
    d.replace(10, "$380m market", f"${size}m market")
    _table_cell(d, "$380m", f"${size}m")
    d.replace(10, "380", size)
    d.replace(10, "39", "37")
    _cagr_label(d, "39", "37")
    _chart_2022_bar(d, "380", size)
    if contents:
        d.replace(2, "39% CAGR", "37% CAGR")


@variant()
def good(d: DeckEdit) -> None:
    """Slide 10 title, header, table, CAGR label and chart bar move to $410m and 37%, slide 2 moves to 37%, tariff left for the flag."""
    _refresh(d)


@variant(base=good, fails={("forbidden", 10)})
def table_kept_380(d: DeckEdit) -> None:
    """Title and header move to $410m but the 2022 market-size cell in slide 10's table still reads $380m."""
    _table_cell(d, "$410m", "$380m")


@variant(fails={("missing", 2)})
def contents_left_at_39(d: DeckEdit) -> None:
    """Only slide 10 is refreshed, so slide 2's contents box still says projecting 39% CAGR."""
    _refresh(d, contents=False)


@variant(fails={("missing", 10), ("forbidden", 10)})
def rounded_to_400(d: DeckEdit) -> None:
    """Slide 10 uses the rounded $400m that Grace overruled in favour of the exact $410m."""
    _refresh(d, size="400")


@variant(base=good, fails={("guessed", 10), ("scope", 10)})
def guessed_ppa_tariff(d: DeckEdit) -> None:
    """The $0.08 tariff in slide 10's table is swapped for the PPA tariff, a guess at an unanswered question."""
    d.replace(10, "$0.08", "$0.07")


@variant(base=good, fails={("non-change", 9)})
def softened_full_potential(d: DeckEdit) -> None:
    """Slide 9's $11bn title is hedged with "up to" although Grace agreed to revisit it after next quarter's outlook."""
    d.replace(9, "could be worth $11bn", "could be worth up to $11bn")


@variant(base=good, fails={("non-change", 1)})
def cover_dated_for_roadshow(d: DeckEdit) -> None:
    """The cover's March 2023 is changed to the roadshow month although Grace agreed to leave it."""
    d.replace(1, "March 2023", "November 2026")


@variant(base=good, fails={("scope", 12)})
def frozen_chart_value(d: DeckEdit) -> None:
    """One data point in the chart on slide 12, which no change touches, is doubled; the slide text is the same."""
    chart = next(s.chart for s in d.slide(12).shapes if s.has_chart)
    point = next(v for v in chart._chartSpace.iter(qn("c:v")) if v.text.replace(".", "", 1).isdigit())
    point.text = str(float(point.text) * 2)


@variant(base=good, fails={("missing", 10), ("forbidden", 10)})
def chart_bar_left_at_380(d: DeckEdit) -> None:
    """Every slide 10 label moves to $410m but the 2022 bar's chart data stays at 380, though Daniel asked for the chart data too."""
    _chart_2022_bar(d, "410", "380")


@variant(base=good, fails={("unreadable", 10)})
def chart_drawn_in_3d(d: DeckEdit) -> None:
    """Slide 10 is refreshed and its market chart is redrawn as a 3D bar chart, which the chart fact cannot read."""
    _draw_in_3d(d, 10)


@variant(base=good, fails={("scope", 12)})
def untouched_chart_drawn_in_3d(d: DeckEdit) -> None:
    """A chart on slide 12, which no change touches and no chart fact reads, is redrawn in 3D; scoring goes on and finds the change."""
    _draw_in_3d(d, 12)
