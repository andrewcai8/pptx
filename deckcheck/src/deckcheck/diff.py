from __future__ import annotations

import difflib
from dataclasses import dataclass
from typing import Literal

from deckcheck.model import Deck, Slide

Status = Literal["unchanged", "changed", "added", "removed"]


@dataclass(frozen=True)
class SlideDiff:
    slide: int
    status: Status
    title: str
    diff: str


@dataclass(frozen=True)
class DeckDiff:
    old: str
    new: str
    old_sha256: str
    new_sha256: str
    slides: tuple[SlideDiff, ...]


def outline(slide: Slide) -> list[str]:
    return [f"{shape.name}: {p.text.strip()}" for shape in slide.shapes for p in shape.paragraphs]


def diff_decks(old: Deck, new: Deck) -> DeckDiff:
    count = max(len(old.slides), len(new.slides))
    slides = []
    for i in range(count):
        a = old.slides[i] if i < len(old.slides) else None
        b = new.slides[i] if i < len(new.slides) else None
        a_lines, b_lines = outline(a) if a else [], outline(b) if b else []
        text = "\n".join(
            difflib.unified_diff(
                a_lines, b_lines, f"old/slide-{i + 1}", f"new/slide-{i + 1}", lineterm=""
            )
        )
        status: Status = "added" if a is None else "removed" if b is None else "changed" if text else "unchanged"
        title = b.title if b else a.title if a else ""
        slides.append(SlideDiff(i + 1, status, title, text))
    return DeckDiff(old.path, new.path, old.sha256, new.sha256, tuple(slides))
