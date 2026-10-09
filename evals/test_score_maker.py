import json
from collections.abc import Callable
from pathlib import Path

import pytest

import score_maker
from scenario import ROOT

SCENARIOS = (
    "fmcg-diagnostic-timeline",
    "insurance-workshop-prep",
    "rcc-flexibility-wording",
    "retail-impact-title",
    "solar-market-refresh",
)
INSURANCE_QUESTION = (
    "Which regulator slide should be punchier (10 the obstacles map, 16 the Singapore example, "
    "or 18 the closing call to cooperate), and in what direction?"
)


@pytest.fixture(autouse=True)
def at_repo_root(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.chdir(ROOT)


def write_run(run_dir: Path, name: str, edit: Callable[[dict], None] = lambda doc: None) -> None:
    doc = json.loads((ROOT / "evals" / name / "changeset.json").read_text())
    edit(doc)
    (run_dir / name).mkdir(parents=True, exist_ok=True)
    (run_dir / name / "changeset.json").write_text(json.dumps(doc))


def run(argv: list, capsys: pytest.CaptureFixture[str]) -> tuple[int, str]:
    code = score_maker.main([str(a) for a in argv])
    return code, capsys.readouterr().out


def test_the_five_fixtures_pass_their_scenarios(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    for name in SCENARIOS:
        write_run(tmp_path, name)

    assert run([tmp_path], capsys) == (
        0,
        "fmcg-diagnostic-timeline: SCENARIO PASS fmcg-diagnostic-timeline (26 checks, 7 intent checks deferred)\n"
        "  flags (reported, not scored): raised none; missing none; 0 unmatched\n"
        "  refs (reported, not scored): 16 of 16 quote their turn\n"
        "insurance-workshop-prep: SCENARIO PASS insurance-workshop-prep (26 checks, 5 intent checks deferred)\n"
        "  flags (reported, not scored): raised n2; missing none; 0 unmatched\n"
        "  refs (reported, not scored): 21 of 21 quote their turn\n"
        "rcc-flexibility-wording: SCENARIO PASS rcc-flexibility-wording (28 checks, 3 intent checks deferred)\n"
        "  flags (reported, not scored): raised none; missing none; 0 unmatched\n"
        "  refs (reported, not scored): 9 of 9 quote their turn\n"
        "retail-impact-title: SCENARIO PASS retail-impact-title (36 checks, 5 intent checks deferred)\n"
        "  flags (reported, not scored): raised none; missing none; 0 unmatched\n"
        "  refs (reported, not scored): 9 of 9 quote their turn\n"
        "solar-market-refresh: SCENARIO PASS solar-market-refresh (41 checks, 7 intent checks deferred)\n"
        "  flags (reported, not scored): raised n1; missing none; 0 unmatched\n"
        "  refs (reported, not scored): 24 of 24 quote their turn\n"
        "MAKER 5 of 5 SCENARIO PASS\n",
    )


def test_the_insurance_flag_is_written_with_source_slide_numbers(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    write_run(tmp_path, "insurance-workshop-prep")

    code, out = run([tmp_path, "insurance-workshop-prep"], capsys)

    assert (code, out.splitlines()[1]) == (0, "  flags (reported, not scored): raised n2; missing none; 0 unmatched")
    assert json.loads((tmp_path / "insurance-workshop-prep" / "flags.json").read_text()) == [
        {"question": INSURANCE_QUESTION, "said": ["00:08:05", "00:08:20", "00:08:38"], "slides": [10, 16, 18]}
    ]


def test_a_changeset_without_flags_overwrites_a_stale_flags_file(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    write_run(tmp_path, "insurance-workshop-prep")
    run([tmp_path, "insurance-workshop-prep"], capsys)
    write_run(tmp_path, "insurance-workshop-prep", lambda doc: doc.update(flags=[]))

    code, out = run([tmp_path, "insurance-workshop-prep"], capsys)

    assert (code, out.splitlines()[1]) == (0, "  flags (reported, not scored): raised none; missing n2; 0 unmatched")
    assert json.loads((tmp_path / "insurance-workshop-prep" / "flags.json").read_text()) == []


def bad_shape(doc: dict) -> None:
    next(c for c in doc["changes"] if c["id"] == "flexibility-bullet")["op"]["shape"] = 999


def test_a_bad_shape_id_is_an_invalid_changeset(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    write_run(tmp_path, "rcc-flexibility-wording", bad_shape)

    assert run([tmp_path, "rcc-flexibility-wording"], capsys) == (
        1,
        "rcc-flexibility-wording: CHANGESET INVALID: flexibility-bullet op.shape: slide 4 (id 10586) has no shape 999; "
        "shapes with text: 2 'Title 1' ('Establish principles to guide…'), 13 'ee4pContent2' ('Approach to defining principl…'), "
        "12 'ee4pContent1' ('Senior leaders Set overall di…'), 15 'ee4pHeader2' ('Decisions to return employees…'), "
        "14 'ee4pHeader1' ('Expectations of both senior l…'), 39 'Oval 20' ('1'), 24 'NavigationText' ('HQ re-opening | Set the found…')\n"
        "MAKER 0 of 1 SCENARIO PASS\n",
    )
    assert not (tmp_path / "rcc-flexibility-wording" / "executed.pptx").exists()


def test_a_scenario_the_maker_never_wrote_fails(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    write_run(tmp_path, "retail-impact-title")

    assert run([tmp_path, "retail-impact-title", "rcc-flexibility-wording"], capsys) == (
        1,
        "retail-impact-title: SCENARIO PASS retail-impact-title (36 checks, 5 intent checks deferred)\n"
        "  flags (reported, not scored): raised none; missing none; 0 unmatched\n"
        "  refs (reported, not scored): 9 of 9 quote their turn\n"
        f"rcc-flexibility-wording: CHANGESET MISSING {tmp_path}/rcc-flexibility-wording/changeset.json\n"
        "MAKER 1 of 2 SCENARIO PASS\n",
    )


def misquote(doc: dict) -> None:
    doc["asks"][0]["refs"][0]["quote"] = "That forty is old."
    doc["held"][0]["refs"][0]["speaker"] = "Katja Mlakar"
    doc["held"][1]["refs"][0]["t"] = "00:00:01"


def test_a_ref_that_does_not_quote_its_turn_is_listed(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    write_run(tmp_path, "insurance-workshop-prep", misquote)

    code, out = run([tmp_path, "insurance-workshop-prep"], capsys)

    assert (code, out.splitlines()[2]) == (
        0,
        "  refs (reported, not scored): 18 of 21 quote their turn; "
        "survey-share 00:01:16: 'That forty is old.' is not in the turn; "
        "oversight-office 00:07:03: Mojca Zupan speaks this turn, not Katja Mlakar; "
        "role-numbers 00:00:01: no turn at this time",
    )


def test_an_unknown_scenario_or_run_dir_is_a_bad_argument(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    assert score_maker.main([str(tmp_path), "no-such-scenario"]) == 2
    assert score_maker.main([str(tmp_path / "absent")]) == 2
    assert capsys.readouterr().err == (
        f"error: no public scenario no-such-scenario in {ROOT / 'evals'}\n"
        f"error: no run directory {tmp_path / 'absent'}\n"
    )
