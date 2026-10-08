"""Facts a slide states, compared by meaning: money, percents and counts by value, chart data by number, concepts by wording."""

from __future__ import annotations

import re
import unicodedata
from collections.abc import Iterable
from dataclasses import dataclass, field, replace
from decimal import Decimal, InvalidOperation
from functools import cache

WHITESPACE = re.compile(r"\s+")
QUOTES = str.maketrans({"‘": "'", "’": "'", "“": '"', "”": '"'})
NUMBER = r"\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+(?:\.\d+)?"
CURRENCIES = {"$": "USD", "us$": "USD", "usd": "USD", "dollar": "USD", "dollars": "USD", "€": "EUR", "eur": "EUR", "euro": "EUR", "euros": "EUR", "£": "GBP", "gbp": "GBP"}
SCALES = {"k": 3, "thousand": 3, "m": 6, "mm": 6, "mn": 6, "million": 6, "b": 9, "bn": 9, "billion": 9}
NUMBER_WORDS = {
    w: n
    for n, w in enumerate(
        "zero one two three four five six seven eight nine ten eleven twelve thirteen fourteen fifteen sixteen seventeen eighteen nineteen twenty".split()
    )
}

START = r"(?<!\w)(?<!\d[.,])"
CURRENCY = r"us\$|usd|eur|gbp|[$€£]"
SCALE = r"thousand|million|billion|mm|mn|bn|k|m|b"
# A currency either leads the number ($100m, US$ 100 million) or follows it (100m$, 100 million dollars). A trailing
# symbol must not lead a number of its own, so "top 5 $100M deals" is $100M and not $5.
MONEY = re.compile(
    rf"(?<![\w$€£])(?P<currency>{CURRENCY})\s?(?P<number>{NUMBER})(?:\s?(?P<scale>{SCALE})(?![a-z]))?(?![.,]?\d)"
    rf"|{START}(?<![$€£])(?P<number_after>{NUMBER})\s?(?:(?P<scale_after>{SCALE})\s?)?(?P<currency_after>{CURRENCY}|dollars?|euros?)(?![a-z]|\s?\d)"
)
PERCENT_UNIT = r"\s?(?:%|per ?cent(?![a-z]))"
# A minus sign makes a different value: -48% is not 48%. A dash between two numbers is a range, and a range states both.
PERCENT = re.compile(rf"{START}(?P<sign>[-−+](?=\d))?(?P<number>{NUMBER}){PERCENT_UNIT}")
PERCENT_RANGE = re.compile(rf"{START}(?P<low>{NUMBER})(?:{PERCENT_UNIT})?\s?[-–—]\s?(?P<high>{NUMBER}){PERCENT_UNIT}")
NUMBER_WORD = rf"(?:{'|'.join(NUMBER_WORDS)})(?![a-z])"
COUNT_WORD = rf"\d+|{'|'.join(NUMBER_WORDS)}"
COUNT_SPEC = re.compile(rf"(?P<number>{COUNT_WORD})[\s-](?P<unit>[a-z]+?)s?")
# Up to three words may sit between a count and its noun ("all 10 of the levers"), but not another number, and not a
# plural that already ends the phrase, so "6 days and then weeks" is not 6 weeks.
COUNT_GAP = rf"(?:(?!{NUMBER_WORD})(?![a-z-]*[^s]s[\s-])[a-z][a-z-]*[\s-]){{0,3}}"


def normalize(text: str) -> str:
    return WHITESPACE.sub(" ", unicodedata.normalize("NFC", text).translate(QUOTES)).strip().casefold()


@dataclass(frozen=True)
class Money:
    amount: Decimal
    currency: str
    text: str = field(compare=False)


@dataclass(frozen=True)
class Percent:
    value: Decimal
    text: str = field(compare=False)


@dataclass(frozen=True)
class Count:
    value: int
    unit: str
    text: str = field(compare=False)


@dataclass(frozen=True)
class ChartValue:
    value: Decimal
    text: str = field(compare=False)


@dataclass(frozen=True)
class Words:
    alternatives: tuple[str, ...]


Value = Money | Percent | Count | ChartValue | Words
KINDS = ("text", "money", "percent", "count", "chart")
EXAMPLES = {"money": "$410m", "percent": "37%", "count": "6 weeks"}


def parse(kind: str, raw: object) -> Value:
    """Read a fact from expected.yaml. A quantity must be the whole string, so a typo cannot silently become a different fact."""
    if kind == "text":
        alternatives = (raw,) if isinstance(raw, str) else tuple(raw) if isinstance(raw, list) else ()
        if not alternatives or not all(isinstance(a, str) and a.strip() for a in alternatives):
            raise ValueError("expected a string or a non-empty list of strings")
        return Words(alternatives)
    if kind == "chart":
        try:
            return ChartValue(Decimal(str(raw)), str(raw))
        except InvalidOperation:
            raise ValueError(f"{raw!r} is not a number") from None
    if not isinstance(raw, str):
        raise ValueError(f"expected a quoted {kind} such as {EXAMPLES[kind]!r}")
    text = normalize(raw)
    if kind == "count":
        m = COUNT_SPEC.fullmatch(text)
        if m is None:
            raise ValueError(f"{raw!r} is not a count such as {EXAMPLES[kind]!r}")
        return Count(_count(m["number"]), m["unit"], raw)
    found = [q for q in _quantities(kind, text) if q[1] == (0, len(text))]
    if len(found) != 1:
        raise ValueError(f"{raw!r} is not one {kind} such as {EXAMPLES[kind]!r}")
    return replace(found[0][0], text=raw)


def match(value: Money | Percent | Count | Words, text: str) -> str | None:
    """The first stretch of text that states the value, or None. A chart value is matched against chart data by in_chart."""
    text = normalize(text)
    match value:
        case Words(alternatives):
            return next((m[0] for alt in alternatives if (m := _phrase(normalize(alt)).search(text))), None)
        case Count(n, unit):
            pattern = re.compile(rf"{START}(?P<number>{COUNT_WORD})[\s-]{COUNT_GAP}{re.escape(unit)}s?(?![a-z])")
            return next((m[0] for m in pattern.finditer(text) if _count(m["number"]) == n), None)
        case Money() | Percent():
            kind = "money" if isinstance(value, Money) else "percent"
            return next((text[a:b] for found, (a, b) in _quantities(kind, text) if found == value), None)


def in_chart(value: ChartValue, numbers: Iterable[Decimal]) -> str | None:
    return f"chart value {value.text}" if value.value in set(numbers) else None


def describe(value: Value) -> str:
    match value:
        case Words(alternatives):
            return " or ".join(repr(a) for a in alternatives)
        case ChartValue(_, text):
            return f"chart value {text}"
        case _:
            return repr(value.text)


def to_json(value: Value) -> dict[str, object]:
    match value:
        case Words(alternatives):
            return {"text": list(alternatives)}
        case Money(amount, currency, text):
            return {"money": text, "amount": f"{amount:f}", "currency": currency}
        case Percent(v, text):
            return {"percent": text, "value": f"{v:f}"}
        case Count(n, unit, text):
            return {"count": text, "value": n, "unit": unit}
        case ChartValue(v, _):
            return {"chart": f"{v:f}"}


def _quantities(kind: str, text: str) -> list[tuple[Money | Percent, tuple[int, int]]]:
    if kind == "money":
        return [
            (
                Money(
                    _number(m["number"] or m["number_after"]).scaleb(SCALES.get(m["scale"] or m["scale_after"] or "", 0)),
                    CURRENCIES[m["currency"] or m["currency_after"]],
                    m[0],
                ),
                m.span(),
            )
            for m in MONEY.finditer(text)
        ]
    ranges = list(PERCENT_RANGE.finditer(text))
    found = [(Percent(_number(m[end]), m[0]), m.span()) for m in ranges for end in ("low", "high")]
    found += [
        (Percent(-_number(m["number"]) if m["sign"] in ("-", "−") else _number(m["number"]), m[0]), m.span())
        for m in PERCENT.finditer(text)
        if not any(r.start() <= m.start() < r.end() for r in ranges)
    ]
    return sorted(found, key=lambda q: q[1])


def _number(digits: str) -> Decimal:
    return Decimal(digits.replace(",", ""))


def _count(word: str) -> int:
    return NUMBER_WORDS[word] if word in NUMBER_WORDS else int(word)


# A phrase cannot start or end inside a longer word or number. A # in it stands for any count, in digits or in words
# up to twenty, so "# consultants" matches "2 consultants" and "two consultants".
@cache
def _phrase(fact: str) -> re.Pattern[str]:
    body = rf"(?:{COUNT_WORD})".join(re.escape(part) for part in fact.split("#"))
    head = (r"(?<![^\W_])" if fact[0].isalnum() or fact[0] == "#" else "") + (r"(?<!\d[.,])" if fact[0].isdigit() or fact[0] == "#" else "")
    tail = (r"(?![^\W_])" if fact[-1].isalnum() or fact[-1] == "#" else "") + (r"(?![.,]\d)" if fact[-1].isdigit() or fact[-1] == "#" else "")
    return re.compile(head + body + tail)
