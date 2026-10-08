"""Facts a slide states, compared by meaning: money, percents and counts by value, chart data by number, concepts by wording."""

from __future__ import annotations

import re
import unicodedata
from collections.abc import Iterable
from dataclasses import dataclass, field, replace
from decimal import Decimal, InvalidOperation

WHITESPACE = re.compile(r"\s+")
QUOTES = str.maketrans({"‘": "'", "’": "'", "“": '"', "”": '"'})
NUMBER = r"\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+(?:\.\d+)?"
CURRENCIES = {"$": "USD", "us$": "USD", "usd": "USD", "€": "EUR", "eur": "EUR", "£": "GBP", "gbp": "GBP"}
SCALES = {"k": 3, "thousand": 3, "m": 6, "mn": 6, "million": 6, "b": 9, "bn": 9, "billion": 9}
NUMBER_WORDS = {
    w: n
    for n, w in enumerate(
        "zero one two three four five six seven eight nine ten eleven twelve thirteen fourteen fifteen sixteen seventeen eighteen nineteen twenty".split()
    )
}

MONEY = re.compile(
    rf"(?<![\w$€£])(?P<currency>us\$|usd|eur|gbp|[$€£])\s?(?P<number>{NUMBER})"
    rf"(?:\s?(?P<scale>thousand|million|billion|mn|bn|k|m|b)(?![a-z]))?(?![.,]?\d)"
)
START = r"(?<!\w)(?<!\d[.,])"
PERCENT = re.compile(rf"{START}(?P<number>{NUMBER})\s?(?:%|per ?cent(?![a-z]))")
COUNT_WORD = rf"\d+|{'|'.join(NUMBER_WORDS)}"
COUNT_SPEC = re.compile(rf"(?P<number>{COUNT_WORD})[\s-](?P<unit>[a-z]+?)s?")


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
    if not found:
        raise ValueError(f"{raw!r} is not one {kind} such as {EXAMPLES[kind]!r}")
    return replace(found[0][0], text=raw)


def match(value: Money | Percent | Count | Words, text: str) -> str | None:
    """The first stretch of text that states the value, or None. A chart value is matched against chart data by in_chart."""
    text = normalize(text)
    match value:
        case Words(alternatives):
            return next((alt for alt in alternatives if _contains(normalize(alt), text)), None)
        case Count(n, unit):
            pattern = re.compile(rf"{START}(?P<number>{COUNT_WORD})[\s-](?:[a-z][a-z-]*[\s-])?{re.escape(unit)}s?(?![a-z])")
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
            (Money(_number(m["number"]).scaleb(SCALES.get(m["scale"] or "", 0)), CURRENCIES[m["currency"]], m[0]), m.span())
            for m in MONEY.finditer(text)
        ]
    return [(Percent(_number(m["number"]), m[0]), m.span()) for m in PERCENT.finditer(text)]


def _number(digits: str) -> Decimal:
    return Decimal(digits.replace(",", ""))


def _count(word: str) -> int:
    return NUMBER_WORDS[word] if word in NUMBER_WORDS else int(word)


def _joined(a: str, b: str, c: str) -> bool:
    return a.isalnum() and b.isalnum() or a.isdigit() and b in ".," and c.isdigit()


def _contains(fact: str, text: str) -> bool:
    text = f"  {text}  "
    start = text.find(fact)
    while start >= 0:
        end = start + len(fact)
        if not _joined(fact[0], text[start - 1], text[start - 2]) and not _joined(fact[-1], text[end], text[end + 1]):
            return True
        start = text.find(fact, start + 1)
    return False
