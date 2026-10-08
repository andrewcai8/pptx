from __future__ import annotations

import math
import re
from collections.abc import Callable, Iterable, Iterator
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import yaml

from deckcheck.model import Deck, Run, Shape, Slide, Violation

Rule = Callable[[Deck, dict[str, Any]], Iterable[Violation]]
Param = Callable[[Any], Any]

RULES: dict[str, Rule] = {}
PARAMS: dict[str, dict[str, Param]] = {}


class ConfigError(Exception):
    pass


@dataclass(frozen=True)
class RuleSet:
    path: Path
    params: dict[str, dict[str, Any]]


def rule(rule_id: str, **params: Param) -> Callable[[Rule], Rule]:
    def register(fn: Rule) -> Rule:
        RULES[rule_id] = fn
        PARAMS[rule_id] = params
        return fn

    return register


def _int(v: Any) -> int:
    if type(v) is not int:
        raise ValueError("expected an integer")
    return v


def _number(v: Any) -> float:
    if type(v) not in (int, float):
        raise ValueError("expected a number")
    return float(v)


def _str(v: Any) -> str:
    if not isinstance(v, str) or not v:
        raise ValueError("expected a non-empty string")
    return v


def _str_list(v: Any) -> tuple[str, ...]:
    if not isinstance(v, list) or not all(isinstance(x, str) for x in v):
        raise ValueError("expected a list of strings")
    return tuple(v)


def _regex_list(v: Any) -> tuple[re.Pattern[str], ...]:
    try:
        return tuple(re.compile(p, re.IGNORECASE) for p in _str_list(v))
    except re.error as e:
        raise ValueError(f"invalid regex: {e}") from e


def load_rules(path: Path) -> RuleSet:
    try:
        doc = yaml.safe_load(path.read_text())
    except (OSError, yaml.YAMLError) as e:
        raise ConfigError(f"cannot read rules {path}: {e}") from e
    if not isinstance(doc, dict) or not isinstance(doc.get("rules"), dict):
        raise ConfigError(f"{path}: expected a top-level `rules:` mapping")
    parsed: dict[str, dict[str, Any]] = {}
    for rule_id, raw in doc["rules"].items():
        if rule_id not in RULES:
            raise ConfigError(f"{path}: unknown rule id {rule_id!r}")
        raw = {} if raw is None else raw
        if not isinstance(raw, dict):
            raise ConfigError(f"{path}: {rule_id}: params must be a mapping")
        spec = PARAMS[rule_id]
        if extra := sorted(set(raw) - set(spec)):
            raise ConfigError(f"{path}: {rule_id}: unknown params {', '.join(extra)}")
        if missing := sorted(set(spec) - set(raw)):
            raise ConfigError(f"{path}: {rule_id}: missing params {', '.join(missing)}")
        try:
            parsed[rule_id] = {name: spec[name](raw[name]) for name in spec}
        except ValueError as e:
            raise ConfigError(f"{path}: {rule_id}: {e}") from e
    return RuleSet(path=path, params=parsed)


def run_rules(deck: Deck, rules: RuleSet) -> list[Violation]:
    found = [v for rule_id, params in rules.params.items() for v in RULES[rule_id](deck, params)]
    return sorted(found, key=lambda v: (v.slide, v.rule, v.shape or ""))


def slide_fonts(slide: Slide) -> set[str]:
    return {r.font for s in slide.shapes for p in s.paragraphs for r in p.runs}


@rule("max-fonts-per-slide", max=_int)
def max_fonts_per_slide(deck: Deck, params: dict[str, Any]) -> Iterator[Violation]:
    for slide in deck.slides:
        fonts = sorted(slide_fonts(slide))
        if len(fonts) > params["max"]:
            yield Violation(
                "max-fonts-per-slide",
                slide.index,
                None,
                f"{len(fonts)} fonts on slide, max {params['max']}",
                ", ".join(fonts),
                target=slide,
            )


@rule("no-bullet-end-punctuation", chars=_str)
def no_bullet_end_punctuation(deck: Deck, params: dict[str, Any]) -> Iterator[Violation]:
    for slide in deck.slides:
        for shape in slide.shapes:
            for p in shape.paragraphs:
                text = p.text.strip()
                if p.is_bullet and text[-1] in params["chars"]:
                    yield Violation(
                        "no-bullet-end-punctuation",
                        slide.index,
                        shape.name,
                        f"bullet ends with {text[-1]!r}",
                        text,
                        target=p,
                    )


# PowerPoint names a copied layout "1_Title Slide".
COPY_PREFIX = re.compile(r"^\d+_")


@rule("slide-has-title", exempt_layouts=_str_list)
def slide_has_title(deck: Deck, params: dict[str, Any]) -> Iterator[Violation]:
    for slide in deck.slides:
        layout = COPY_PREFIX.sub("", slide.layout_name)
        if layout not in params["exempt_layouts"] and not slide.title:
            yield Violation(
                "slide-has-title", slide.index, None, "slide has no title", f"layout {slide.layout_name}", target=slide
            )


@rule("title-max-chars", max=_int)
def title_max_chars(deck: Deck, params: dict[str, Any]) -> Iterator[Violation]:
    for slide in deck.slides:
        if len(slide.headline) > params["max"] and (shape := slide.title_shape):
            yield Violation(
                "title-max-chars",
                slide.index,
                shape.name,
                f"title is {len(slide.headline)} chars, max {params['max']}",
                slide.headline,
                target=shape,
            )


def small_runs(shape: Shape, min_pt: float) -> list[Run]:
    return [r for p in shape.paragraphs for r in p.runs if r.size_pt is not None and r.size_pt < min_pt]


@rule("min-font-size", min_pt=_number)
def min_font_size(deck: Deck, params: dict[str, Any]) -> Iterator[Violation]:
    for slide in deck.slides:
        for shape in slide.shapes:
            small = small_runs(shape, params["min_pt"])
            if small:
                worst = min(small, key=lambda r: r.size_pt or 0)
                yield Violation(
                    "min-font-size",
                    slide.index,
                    shape.name,
                    f"{worst.size_pt:g}pt text, min {params['min_pt']:g}pt",
                    worst.text,
                    target=shape,
                )


@rule("no-placeholder-text", patterns=_regex_list)
def no_placeholder_text(deck: Deck, params: dict[str, Any]) -> Iterator[Violation]:
    for slide in deck.slides:
        for shape in slide.shapes:
            for p in shape.paragraphs:
                match = next((m for pat in params["patterns"] if (m := pat.search(p.text))), None)
                if match:
                    yield Violation(
                        "no-placeholder-text",
                        slide.index,
                        shape.name,
                        f"placeholder text {match.group(0)!r}",
                        p.text.strip(),
                        target=p,
                    )


EMU_PER_PT = 12700


# The axis-aligned box a shape covers once rotated about its centre.
def visual_box(s: Shape) -> tuple[int, int, int, int]:
    angle = math.radians(s.rotation)
    cos, sin = abs(math.cos(angle)), abs(math.sin(angle))
    width, height = s.width * cos + s.height * sin, s.width * sin + s.height * cos
    cx, cy = s.left + s.width / 2, s.top + s.height / 2
    return round(cx - width / 2), round(cy - height / 2), round(cx + width / 2), round(cy + height / 2)


@rule("within-slide-bounds", tolerance_pt=_number)
def within_slide_bounds(deck: Deck, params: dict[str, Any]) -> Iterator[Violation]:
    tolerance = params["tolerance_pt"] * EMU_PER_PT
    for slide in deck.slides:
        for s in slide.shapes:
            if not s.paragraphs and s.kind != "chart":
                continue
            left, top, right, bottom = visual_box(s)
            if right <= 0 or bottom <= 0 or left >= deck.slide_width or top >= deck.slide_height:
                continue
            overhang = max(-left, -top, right - deck.slide_width, bottom - deck.slide_height)
            if overhang > tolerance:
                yield Violation(
                    "within-slide-bounds",
                    slide.index,
                    s.name,
                    f"text shape extends {overhang / EMU_PER_PT:.0f}pt past the slide edge, "
                    f"tolerance {params['tolerance_pt']:g}pt",
                    f"box ({left}, {top}, {right}, {bottom}) vs slide ({deck.slide_width}, {deck.slide_height}) EMU",
                    target=s,
                )


# A table holds data when a cell has a percent, a currency symbol, or a number like 3.5 or 1,200.
NUMBER_LIKE = re.compile(r"[%$€£¥]|\d[.,]\d")


def _is_data(shape: Shape) -> bool:
    return shape.kind == "chart" or (
        shape.kind == "table" and any(NUMBER_LIKE.search(p.text) for p in shape.paragraphs)
    )


@rule("source-on-data-slides", prefix=_str)
def source_on_data_slides(deck: Deck, params: dict[str, Any]) -> Iterator[Violation]:
    for slide in deck.slides:
        data = [s for s in slide.shapes if _is_data(s)]
        label = re.compile(rf"\s{re.escape(params['prefix'])}s?\s*:")
        has_source = any(
            line.startswith(params["prefix"]) or label.search(line)
            for s in slide.shapes
            if s.kind != "table"
            for p in s.paragraphs
            for line in p.lines
        )
        if data and not has_source:
            yield Violation(
                "source-on-data-slides",
                slide.index,
                data[0].name,
                f"{data[0].kind} slide has no line starting with {params['prefix']!r}",
                ", ".join(s.name for s in data),
                target=data[0],
            )
