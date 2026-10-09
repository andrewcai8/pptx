import json
import re
from collections.abc import Iterator
from pathlib import Path

import pytest
import yaml

from deckcheck.meeting import COMMANDS, DOCS, SKILL
from scenario import EVALS, ROOT, public

FACT_KEYS = ("money", "percent", "count", "text")
SPEAKER = re.compile(r"^\[\d\d:\d\d:\d\d\] ([^(]+?) \(", re.M)
TURN = re.compile(r"^\[(\d\d:\d\d:\d\d)\]", re.M)
FENCED = re.compile(r"^```bash\n(.*?)^```", re.M | re.S)


def facts(node: object) -> Iterator[str]:
    match node:
        case dict():
            for key, value in node.items():
                if key in FACT_KEYS:
                    yield from [value] if isinstance(value, str) else value
                elif key == "chart":
                    yield str(value)
                else:
                    yield from facts(value)
        case list():
            for item in node:
                yield from facts(item)


def quoted(change: dict) -> Iterator[str]:
    op = change["op"]
    yield from (str(op[k]) for k in ("old", "new", "text", "layout") if k in op)
    yield from (p["text"] for p in op.get("paragraphs", ()))
    yield from (str(op[k]) for k in ("slide", "after") if isinstance(op.get(k), int))


def answers(scenario: Path) -> set[str]:
    key = yaml.safe_load((scenario / "expected.yaml").read_text())
    fixture = json.loads((scenario / "changeset.json").read_text())
    found = {f for f in facts(key) if "#" not in f}
    found |= {q for c in fixture["changes"] for q in quoted(c)}
    found |= {str(s) for item in (*fixture["flags"], *fixture["held"]) for s in item.get("slides", ())}
    found |= {item["id"] for key in ("asks", "changes", "flags", "held") for item in fixture[key]}
    transcript = (scenario / "transcript.md").read_text()
    found |= set(SPEAKER.findall(transcript)) | set(TURN.findall(transcript))
    return {a for a in found if len(a) >= 3}


def leaks(text: str, scenarios: list[Path]) -> list[tuple[str, str]]:
    return sorted(
        (s.name, a)
        for s in scenarios
        for a in answers(s)
        if re.search(rf"(?<![\w.$]){re.escape(a)}(?![\w%])", text, re.IGNORECASE)
    )


@pytest.mark.parametrize("doc", DOCS)
def test_no_file_the_maker_reads_holds_an_answer_from_a_public_scenario(doc: str) -> None:
    assert leaks((ROOT / doc).read_text(), public()) == []


def test_a_skill_that_quotes_a_scenario_answer_is_caught() -> None:
    text = 'Replace "ca. 50%" on slide 301 with the 48% that Mojca Zupan gave at 00:05:46, as for slide 2147478638 and cover-date.'

    assert leaks(text, [EVALS / "insurance-workshop-prep", EVALS / "solar-market-refresh"]) == [
        ("insurance-workshop-prep", "301"),
        ("insurance-workshop-prep", "48%"),
        ("insurance-workshop-prep", "50%"),
        ("insurance-workshop-prep", "Mojca Zupan"),
        ("insurance-workshop-prep", "ca. 50%"),
        ("solar-market-refresh", "00:05:46"),
        ("solar-market-refresh", "2147478638"),
        ("solar-market-refresh", "cover-date"),
    ]


def test_the_skill_runs_exactly_the_commands_the_runner_allows() -> None:
    skill = (ROOT / SKILL).read_text()
    runs = {line for block in FENCED.findall(skill) for line in block.splitlines() if line.startswith("uv run")}

    assert runs == set(COMMANDS)
