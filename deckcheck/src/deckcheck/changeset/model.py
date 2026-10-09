from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Annotated, ClassVar, Literal, get_args

from pydantic import AfterValidator, BaseModel, ConfigDict, Field, ValidationError
from pydantic_core import PydanticCustomError

DecisionKind = Literal["keep_new", "keep_old", "edited"]
SLIDE_DECISIONS: tuple[DecisionKind, ...] = ("keep_new", "keep_old")
EDIT_DECISIONS: tuple[DecisionKind, ...] = ("keep_new", "keep_old", "edited")


class Wire(BaseModel):
    model_config = ConfigDict(frozen=True, extra="forbid", strict=True)


SlideId = Annotated[int, Field(ge=1, description="p:sldId/@id of a source slide")]
ShapeId = Annotated[int, Field(ge=0, description="p:cNvPr/@id, unique on its slide, found inside groups too")]
ChangeId = Annotated[str, Field(pattern=r"^[A-Za-z0-9][A-Za-z0-9_.-]*$")]
Index = Annotated[int, Field(ge=0)]
NonEmpty = Annotated[str, Field(min_length=1)]
Number = Annotated[float, Field(allow_inf_nan=False)]
NUMBER_TEXT = re.compile(r"^[+-]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?$")
CONTROL = re.compile(r"[\x00-\x08\x0c-\x1f\x7f-\x9f\ufffe\uffff]")


def _no_control(value: str) -> str:
    if found := CONTROL.search(value):
        raise PydanticCustomError(
            "control_character",
            "holds the control character {char} at index {index}, which a deck cannot hold; remove it",
            {"char": repr(found.group()), "index": found.start()},
        )
    return value


# Tab is kept; "\n" and "\v" mean a paragraph and a line break, so each op decides where they may go.
DeckText = Annotated[str, AfterValidator(_no_control)]


class Ref(Wire):
    t: Annotated[str, Field(pattern=r"^\d{2}:\d{2}:\d{2}$", description="HH:MM:SS of the transcript turn")]
    speaker: NonEmpty
    quote: Annotated[str, Field(min_length=1, description="copied verbatim from the transcript")]


Refs = Annotated[tuple[Ref, ...], Field(min_length=1)]


class Meeting(Wire):
    title: NonEmpty
    date: Annotated[str, Field(pattern=r"^\d{4}-\d{2}-\d{2}$")]


class Source(Wire):
    path: Annotated[str, Field(min_length=1, description="the source deck, relative to the repo root")]
    sha256: Annotated[str, Field(pattern=r"^[0-9a-f]{64}$")]


class Ask(Wire):
    id: ChangeId
    text: NonEmpty
    refs: Refs


class Flag(Wire):
    id: ChangeId
    question: NonEmpty
    slides: tuple[SlideId, ...] = ()
    refs: Refs


class Held(Wire):
    id: ChangeId
    text: NonEmpty
    slides: tuple[SlideId, ...] = ()
    refs: Refs


class Edited(Wire):
    edited: str


Decision = Literal["pending", "keep_new", "keep_old"] | Edited


class Op(Wire):
    structural: ClassVar[bool] = True
    admits: ClassVar[tuple[DecisionKind, ...]] = SLIDE_DECISIONS

    def edit(self, text: str) -> Op:
        raise ValueError(f"{self.kind} admits keep_new or keep_old only")

    def _edited(self, text: str, /, **update: object) -> Op:
        """The op with the reviewer's text, validated like the maker's: model_copy would skip validation."""
        try:
            return self.model_validate({**dict(self), **update})
        except ValidationError as e:
            raise ValueError(f"edited text {text!r}: {e.errors()[0]['msg']}") from None


class ReplaceText(Op):
    """Replace the quote `old`, which occurs exactly once in the shape (or in paragraph `paragraph`), with `new`.
    Text is a:r and a:fld text in order with a:br as "\\v", as python-pptx's paragraph.text reads it."""

    structural: ClassVar[bool] = False
    admits: ClassVar[tuple[DecisionKind, ...]] = EDIT_DECISIONS
    kind: Literal["replace_text"]
    slide: SlideId
    shape: ShapeId
    old: NonEmpty
    new: DeckText
    paragraph: Index | None = Field(
        default=None, description="index among all a:p of the shape, blank ones included; only to tell apart a quote that repeats"
    )

    def edit(self, text: str) -> Op:
        return self._edited(text, new=text)


class InsertParagraph(Op):
    """A new paragraph after source paragraph `after`, styled like it."""

    structural: ClassVar[bool] = False
    admits: ClassVar[tuple[DecisionKind, ...]] = EDIT_DECISIONS
    kind: Literal["insert_paragraph"]
    slide: SlideId
    shape: ShapeId
    after: Index
    text: Annotated[NonEmpty, AfterValidator(_no_control)]

    def edit(self, text: str) -> Op:
        return self._edited(text, text=text)


class SetCell(Op):
    """The whole text of table cell (row, col), paragraphs joined by "\\n". `old` must equal the source cell."""

    admits: ClassVar[tuple[DecisionKind, ...]] = EDIT_DECISIONS
    kind: Literal["set_cell"]
    slide: SlideId
    shape: ShapeId
    row: Index
    col: Index
    old: str
    new: DeckText

    def edit(self, text: str) -> Op:
        return self._edited(text, new=text)


class SetChartValue(Op):
    """Point `point` (c:pt/@idx) of the `series`-th c:ser in document order: the chart cache and its workbook cell."""

    admits: ClassVar[tuple[DecisionKind, ...]] = EDIT_DECISIONS
    kind: Literal["set_chart_value"]
    slide: SlideId
    shape: ShapeId
    series: Index
    point: Index
    old: Number
    new: Number

    def edit(self, text: str) -> Op:
        if not NUMBER_TEXT.match(text.strip()):
            raise ValueError(f"edited value {text!r} is not a number like 1234.5 or 1,234.5")
        return self._edited(text, new=float(text.strip().replace(",", "")))


class AddSlide(Op):
    """A slide on one of the deck's own layouts, after source slide `after` (null: first), placeholders empty."""

    kind: Literal["add_slide"]
    layout: NonEmpty
    after: SlideId | None


class Paragraph(Wire):
    text: DeckText
    level: Annotated[int, Field(ge=0, le=8)] = 0


class FillPlaceholder(Op):
    """The paragraphs of placeholder `shape` (its cNvPr id on the layout) on the slide add_slide change `slide` adds."""

    admits: ClassVar[tuple[DecisionKind, ...]] = EDIT_DECISIONS
    kind: Literal["fill_placeholder"]
    slide: ChangeId
    shape: ShapeId
    paragraphs: Annotated[tuple[Paragraph, ...], Field(min_length=1)]

    def edit(self, text: str) -> Op:
        levels = [p.level for p in self.paragraphs]
        lines = text.split("\n")
        return self._edited(text, paragraphs=tuple({"text": t, "level": levels[min(i, len(levels) - 1)]} for i, t in enumerate(lines)))


class DeleteSlide(Op):
    kind: Literal["delete_slide"]
    slide: SlideId


class MoveSlide(Op):
    """Move source slide `slide` after source slide `after` (null: first), wherever `after` ends up."""

    kind: Literal["move_slide"]
    slide: SlideId
    after: SlideId | None


AnyOp = Annotated[
    ReplaceText | InsertParagraph | SetCell | SetChartValue | AddSlide | FillPlaceholder | DeleteSlide | MoveSlide,
    Field(discriminator="kind"),
]
OP_KINDS: tuple[str, ...] = tuple(get_args(c.model_fields["kind"].annotation)[0] for c in get_args(get_args(AnyOp)[0]))


class Change(Wire):
    id: ChangeId
    ask_id: ChangeId
    rationale: Annotated[str, Field(min_length=1, pattern=r"^[^\n]+$", description="one sentence")]
    refs: Refs
    op: AnyOp
    decision: Decision = Field(default="pending", description="written by the review app; absent means pending")


class ChangeSet(Wire):
    meeting: Meeting
    source: Source
    asks: tuple[Ask, ...]
    changes: tuple[Change, ...]
    flags: tuple[Flag, ...] = ()
    held: tuple[Held, ...] = ()


SlideKey = Annotated[int | str, Field(description="a source slide id, or the id of the add_slide change that adds it")]


class Box(Wire):
    x: int
    y: int
    w: int
    h: int


class ReviewShape(Wire):
    id: ShapeId
    name: str
    box: Box | None = Field(description="EMU on the slide, group transforms applied")


class ReviewSlide(Wire):
    key: SlideKey
    source_index: int | None = Field(description="1-based in the source deck; null for an added slide")
    executed_index: int | None = Field(description="1-based in the executed deck; null for a deleted slide")
    title: str


class ReviewChange(Wire):
    id: ChangeId
    ask_id: ChangeId
    kind: str
    structural: bool
    slide: SlideKey
    source_index: int | None
    executed_index: int | None
    shape: ReviewShape | None
    before: str | None = Field(description="read from the source deck, never from the maker")
    after: str | None
    span: tuple[int, int] | None = Field(description="replace_text: [start, end) of the quote in `before`, in code points")
    depends_on: ChangeId | None
    admits: tuple[DecisionKind, ...]
    decision: Decision
    rationale: str
    refs: tuple[Ref, ...]
    notes: tuple[str, ...]


class Executed(Wire):
    path: str
    sha256: str


class Review(Wire):
    source: Source
    executed: Executed
    meeting: Meeting
    asks: tuple[Ask, ...]
    flags: tuple[Flag, ...]
    held: tuple[Held, ...]
    slides: tuple[ReviewSlide, ...]
    changes: tuple[ReviewChange, ...]


SCHEMA_DIR = Path(__file__).parent


def schemas() -> dict[str, str]:
    return {
        name: json.dumps(model.model_json_schema(), indent=2) + "\n"
        for name, model in (("changeset.schema.json", ChangeSet), ("review.schema.json", Review))
    }


if __name__ == "__main__":
    for name, text in schemas().items():
        (SCHEMA_DIR / name).write_text(text)
