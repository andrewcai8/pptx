from __future__ import annotations

import functools
import hashlib
import http.client
import json
import os
import re
import shutil
import socket
import subprocess
import sys
import threading
import time
import urllib.error
import urllib.request
import zipfile
from collections.abc import Callable, Iterator
from dataclasses import dataclass
from pathlib import Path

import pytest
from test_changeset import ADD, CELL, FILL, REF, REVENUE, build_deck

from deckcheck.changeset.cli import main as changeset_main
from deckcheck.cli import write_atomic
from deckcheck.render import find_fc_match, find_pdftoppm, find_soffice, render
from deckcheck.review.cli import main as review_main
from deckcheck.review.maker import maker_for
from deckcheck.review import meetings
from deckcheck.review.meetings import Locked, Reviews, cascade
from deckcheck.review.server import serve

WEB = Path(__file__).parents[1] / "src/deckcheck/review/web"
INSERT = {"kind": "insert_paragraph", "slide": 256, "shape": 3, "after": 0, "text": "Costs fall"}
DEMO = [("c1", REVENUE), ("c2", INSERT), ("c3", CELL), ("add", ADD), ("fill", {**FILL, "slide": "add"})]
SIMULATED = {"label": "Simulated maker: replays the committed changeset.json", "simulated": True}
PENDING = {"c1": "pending", "c2": "pending", "c3": "pending", "add": "pending", "fill": "pending"}
GATE = threading.Event()


def sha(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def fake_render(deck: Path, out: Path) -> None:
    GATE.wait(10)
    with zipfile.ZipFile(deck) as z:
        count = sum(1 for n in z.namelist() if n.startswith("ppt/slides/slide") and n.endswith(".xml"))
    out.mkdir(parents=True)
    for i in range(1, count + 1):
        (out / f"slide-{i}.png").write_bytes(b"\x89PNG " + f"{out.name} {i}".encode())
    fonts = {"Calibri": {"family": "Carlito", "substituted": False}, "Nonexistent Sans QA": {"family": "DejaVu Sans", "substituted": True}}
    (out / "fonts.json").write_text(json.dumps(fonts))


def write_meeting(name: str, changes: list[tuple[str, dict]], deck: Path = Path("decks/deck.pptx")) -> Path:
    d = Path("evals") / name
    d.mkdir(parents=True)
    (d / "transcript.md").write_text("# Pricing review\n\nDate: 2026-10-01, video call\n\nAna Ruiz: Use the new figures.\n")
    doc = {
        "meeting": {"title": "Pricing review", "date": "2026-10-01"},
        "source": {"path": str(deck), "sha256": sha(deck.read_bytes())},
        "asks": [{"id": "a1", "text": "Refresh the figures.", "refs": [REF]}],
        "changes": [{"id": cid, "ask_id": "a1", "rationale": "The figures moved.", "refs": [REF], "op": op} for cid, op in changes],
    }
    (d / "changeset.json").write_text(json.dumps(doc, indent=1))
    return d / "changeset.json"


@pytest.fixture
def repo(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    monkeypatch.chdir(tmp_path)
    GATE.set()
    (tmp_path / "deckcheck").mkdir()
    (tmp_path / "deckcheck/pyproject.toml").write_text("")
    (tmp_path / "decks").mkdir()
    build_deck(tmp_path / "decks/deck.pptx")
    write_meeting("demo", DEMO)
    return tmp_path


@dataclass(frozen=True)
class Client:
    port: int
    stop: Callable[[], None]

    def call(self, method: str, path: str, body: object = None, headers: dict | None = None) -> tuple[int, object]:
        data = json.dumps(body).encode() if method == "POST" else None
        req = urllib.request.Request(
            f"http://127.0.0.1:{self.port}{path}", data=data, method=method, headers={"Content-Type": "application/json", **(headers or {})}
        )
        try:
            with urllib.request.urlopen(req) as r:
                status, raw, kind = r.status, r.read(), r.headers.get_content_type()
        except urllib.error.HTTPError as e:
            status, raw, kind = e.code, e.read(), e.headers.get_content_type()
        return status, json.loads(raw) if kind == "application/json" else raw

    def get(self, path: str) -> tuple[int, object]:
        return self.call("GET", path)

    def post(self, path: str, body: object = None) -> tuple[int, object]:
        return self.call("POST", path, {} if body is None else body)

    def row(self, mid: str) -> dict:
        return next(m for m in self.get("/api/meetings")[1]["meetings"] if m["id"] == mid)

    def settle(self, mid: str) -> dict:
        deadline = time.monotonic() + 10
        while (row := self.row(mid))["state"]["is"] == "processing":
            assert time.monotonic() < deadline, "processing did not finish in 10 s"
            time.sleep(0.02)
        return row

    def ready(self, mid: str = "evals/demo") -> dict:
        assert self.post(f"/api/meetings/{mid}/process")[0] == 202
        return self.settle(mid)

    def decide(self, mid: str = "evals/demo", **decisions: object) -> tuple[int, object]:
        return self.post(f"/api/meetings/{mid}/decisions", {"decisions": decisions})


@pytest.fixture
def start(repo: Path) -> Iterator:
    clients = []

    def start(make=maker_for, render=fake_render) -> Client:
        reviews = Reviews(render=render, maker_for=make)
        server = serve(reviews, port=0)
        threading.Thread(target=server.serve_forever, daemon=True).start()

        def stop() -> None:
            server.shutdown()
            server.server_close()
            reviews.close()

        clients.append(Client(server.server_address[1], stop))
        return clients[-1]

    yield start
    GATE.set()
    for client in clients:
        client.stop()


@pytest.fixture
def app(start) -> Client:
    return start()


def test_meetings_list_each_meeting_dir_with_its_maker(repo: Path, app: Client) -> None:
    (repo / "private/meetings/x").mkdir(parents=True)
    (repo / "private/meetings/x/transcript.md").write_text("# Client sync\n\nNo date here.\n")
    (repo / "evals/__pycache__").mkdir()
    (repo / "evals/__pycache__/changeset.json").write_text("{}")
    (repo / "evals/notes").mkdir()

    assert app.get("/api/meetings") == (
        200,
        {
            "meetings": [
                {
                    "id": "evals/demo",
                    "origin": "evals",
                    "name": "demo",
                    "title": "Pricing review",
                    "date": "2026-10-01",
                    "maker": SIMULATED,
                    "state": {"is": "new"},
                    "fonts": [],
                },
                {
                    "id": "private/x",
                    "origin": "private",
                    "name": "x",
                    "title": "Client sync",
                    "date": None,
                    "maker": None,
                    "state": {"is": "new"},
                    "fonts": [],
                },
            ]
        },
    )
    assert app.post("/api/meetings/private/x/process") == (409, {"error": "this meeting has no changeset.json; start the server with --maker to process it"})
    assert app.get("/api/meetings/evals/__pycache__") == (404, {"error": "not found"})
    assert app.get("/api/meetings/evals/notes") == (404, {"error": "no meeting evals/notes"})
    assert app.get("/api/meetings/evals/demo.partial") == (404, {"error": "not found"})


def test_process_executes_and_renders_both_decks(repo: Path, app: Client) -> None:
    assert app.ready() == {
        "id": "evals/demo",
        "origin": "evals",
        "name": "demo",
        "title": "Pricing review",
        "date": "2026-10-01",
        "maker": SIMULATED,
        "state": {"is": "ready", "decided": 0, "total": 5, "applied": False},
        "fonts": [{"font": "Nonexistent Sans QA", "family": "DejaVu Sans"}],
    }
    status, view = app.get("/api/meetings/evals/demo")
    executed = (repo / "artifacts/review/evals/demo/executed.pptx").read_bytes()

    assert status == 200
    assert [c["id"] for c in view["review"]["changes"]] == ["c1", "c2", "c3", "add", "fill"]
    assert [c["decision"] for c in view["review"]["changes"]] == ["pending"] * 5
    assert view["review"]["executed"] == {"path": "artifacts/review/evals/demo/executed.pptx", "sha256": sha(executed)}
    assert view["slide"] == {"w": 9144000, "h": 6858000}
    assert view["fonts"] == [{"font": "Nonexistent Sans QA", "family": "DejaVu Sans"}]
    assert view["final"] is None
    assert re.fullmatch(r"/api/meetings/evals/demo/render/old/slide-\{n\}\.png\?v=[0-9a-f]{12}", view["images"]["old"])
    assert view["images"]["new"] == view["images"]["old"].replace("/old/", "/new/")
    assert app.get(view["images"]["old"].replace("{n}", "3")) == (200, b"\x89PNG old 3")
    assert app.get(view["images"]["new"].replace("{n}", "4")) == (200, b"\x89PNG new 4")
    assert app.get("/api/meetings/evals/demo/render/old/slide-4.png") == (404, {"error": "not found"})


def test_processing_again_gives_the_slide_pictures_new_urls(repo: Path, app: Client) -> None:
    app.ready()
    first = app.get("/api/meetings/evals/demo")[1]
    app.post("/api/meetings/evals/demo/process", {"again": True})
    app.settle("evals/demo")
    second = app.get("/api/meetings/evals/demo")[1]

    assert second["review"]["executed"] == first["review"]["executed"]
    assert second["images"]["new"] != first["images"]["new"]


def test_processing_a_ready_meeting_again_needs_asking(repo: Path, app: Client) -> None:
    app.ready()
    app.decide(c1="keep_old")

    assert app.post("/api/meetings/evals/demo/process")[1]["state"] == {"is": "ready", "decided": 1, "total": 5, "applied": False}
    assert app.post("/api/meetings/evals/demo/process", {"again": True})[0] == 202
    assert app.settle("evals/demo")["state"] == {"is": "ready", "decided": 0, "total": 5, "applied": False}


@pytest.fixture
def umask_022() -> Iterator[None]:
    old = os.umask(0o022)
    yield
    os.umask(old)


def test_a_decision_is_written_to_the_working_changeset_only(repo: Path, app: Client, umask_022: None) -> None:
    committed = (repo / "evals/demo/changeset.json").read_bytes()
    app.ready()

    assert app.decide(c1="keep_old", c2={"edited": "Costs fall 3%"}) == (
        200,
        {"decisions": {**PENDING, "c1": "keep_old", "c2": {"edited": "Costs fall 3%"}}, "final": None},
    )
    path = repo / "artifacts/review/evals/demo/changeset.json"
    working = json.loads(path.read_text())
    assert [c.get("decision") for c in working["changes"]] == ["keep_old", {"edited": "Costs fall 3%"}, None, None, None]
    assert oct(path.stat().st_mode & 0o777) == "0o644"
    assert (repo / "evals/demo/changeset.json").read_bytes() == committed
    assert app.decide(c1="pending")[1]["decisions"]["c1"] == "pending"
    assert app.get("/api/meetings/evals/demo")[1]["review"]["changes"][1]["decision"] == {"edited": "Costs fall 3%"}


def test_a_decision_the_engine_refuses_leaves_the_file_alone(repo: Path, app: Client) -> None:
    app.ready()
    working = repo / "artifacts/review/evals/demo/changeset.json"
    before = working.read_bytes()

    assert app.decide(c2={"edited": ""}) == (422, {"problems": [{"where": "c2 decision", "message": "edited text '': String should have at least 1 character"}]})
    assert app.decide(add={"edited": "x"}) == (422, {"problems": [{"where": "add decision", "message": "add_slide admits keep_new or keep_old only"}]})
    assert app.decide(nope="keep_new") == (422, {"problems": [{"where": "nope", "message": "no such change"}]})
    assert app.decide(c1="maybe")[0] == 400
    assert app.post("/api/meetings/evals/demo/decisions", {"decisions": {}, "extra": 1})[0] == 400
    assert working.read_bytes() == before
    assert sorted(p.name for p in working.parent.iterdir()) == ["changeset.json", "executed.pptx", "render"]


def test_dropping_an_added_slide_settles_only_its_undecided_fills(repo: Path, app: Client) -> None:
    app.ready()
    edited = {"edited": "Raise list prices\nHold discounts"}

    assert app.decide(add="keep_old")[1]["decisions"] == {**PENDING, "add": "keep_old", "fill": "keep_old"}
    assert app.decide(add="keep_new")[1]["decisions"] == {**PENDING, "add": "keep_new", "fill": "pending"}
    app.decide(fill=edited)
    assert app.decide(add="keep_old")[1]["decisions"] == {**PENDING, "add": "keep_old", "fill": edited}
    assert app.decide(add="keep_new")[1]["decisions"] == {**PENDING, "add": "keep_new", "fill": edited}
    app.decide(fill="keep_old")
    assert app.decide(add="keep_new")[1]["decisions"] == {**PENDING, "add": "keep_new", "fill": "keep_old"}


def test_apply_reports_a_kept_fill_of_a_dropped_slide_as_dropped(repo: Path, app: Client) -> None:
    app.ready()
    app.decide(c1="keep_new", c2="keep_new", c3="keep_new", fill="keep_new")
    app.decide(add="keep_old")

    status, final = app.post("/api/meetings/evals/demo/apply")

    assert status == 200
    assert (final["kept_new"], final["kept_old"], final["dropped"]) == (["c1", "c2", "c3"], ["add"], [["fill", "add"]])


def test_cascade_judges_a_restore_against_the_decision_on_file() -> None:
    fill = {"id": "fill", "op": {**FILL, "slide": "add"}, "decision": "keep_old"}

    assert cascade([{"id": "add", "op": ADD}, fill], {"add": "keep_old", "fill": "keep_new"}) == {"add": "keep_old", "fill": "keep_new"}
    assert cascade([{"id": "add", "op": ADD, "decision": "keep_old"}, fill], {"add": "keep_new"}) == {"add": "keep_new", "fill": "pending"}
    assert cascade([{"id": "add", "op": ADD, "decision": "keep_new"}, fill], {"add": "keep_new"}) == {"add": "keep_new"}


def test_apply_writes_the_bytes_the_changeset_cli_writes(repo: Path, app: Client) -> None:
    app.ready()
    assert app.post("/api/meetings/evals/demo/apply") == (409, {"error": "decide every change first", "pending": ["c1", "c2", "c3", "add", "fill"]})
    app.decide(c1={"edited": "16%"}, c2="keep_new", c3="keep_old", add="keep_new", fill="keep_new")
    working = repo / "artifacts/review/evals/demo/changeset.json"
    shutil.copyfile(working, repo / "copy.json")
    assert changeset_main(["apply", "copy.json", "--out", "cli.pptx"]) == 0
    cli = (repo / "cli.pptx").read_bytes()

    assert app.post("/api/meetings/evals/demo/apply") == (
        200,
        {
            "path": "artifacts/review/evals/demo/final.pptx",
            "sha256": sha(cli),
            "changeset_sha256": sha(working.read_bytes()),
            "kept_new": ["c2", "add", "fill"],
            "edited": ["c1"],
            "kept_old": ["c3"],
            "dropped": [],
            "download": "/api/meetings/evals/demo/final.pptx",
        },
    )
    assert (repo / "artifacts/review/evals/demo/final.pptx").read_bytes() == cli
    with urllib.request.urlopen(f"http://127.0.0.1:{app.port}/api/meetings/evals/demo/final.pptx") as r:
        assert (r.read(), r.headers["Content-Disposition"]) == (cli, 'attachment; filename="demo-final.pptx"')
    assert app.row("evals/demo")["state"] == {"is": "ready", "decided": 5, "total": 5, "applied": True}
    assert app.decide(c3="keep_old")[1]["final"]["sha256"] == sha(cli)
    assert app.decide(c3="keep_new")[1]["final"] is None
    assert app.get("/api/meetings/evals/demo/final.pptx")[0] == 404
    assert app.row("evals/demo")["state"]["applied"] is False


def test_a_failed_apply_never_leaves_an_older_final_deck_current(repo: Path, app: Client, monkeypatch: pytest.MonkeyPatch) -> None:
    app.ready()
    app.decide(c1="keep_new", c2="keep_new", c3="keep_new", add="keep_new", fill="keep_new")
    app.post("/api/meetings/evals/demo/apply")
    app.decide(c3="keep_old")

    def crash(path: Path, data: bytes) -> None:
        if path.name == "applied.json":
            raise OSError("disk full")
        write_atomic(path, data)

    monkeypatch.setattr(meetings, "write_atomic", crash)
    assert app.post("/api/meetings/evals/demo/apply") == (500, {"error": "internal error: disk full"})

    assert app.decide(c3="keep_new")[1]["final"] is None
    assert app.get("/api/meetings/evals/demo/final.pptx") == (404, {"error": "no final deck for the current decisions; apply them first"})


def test_decisions_wait_while_the_meeting_is_processed(repo: Path, app: Client) -> None:
    app.ready()
    GATE.clear()
    app.post("/api/meetings/evals/demo/process", {"again": True})

    assert app.decide(c1="keep_new") == (409, {"error": "this meeting is being processed; wait until it is ready"})
    assert app.post("/api/meetings/evals/demo/apply")[0] == 409
    assert app.get("/api/meetings/evals/demo")[1]["state"]["is"] == "processing"
    deadline = time.monotonic() + 10
    while app.row("evals/demo")["state"] != {"is": "processing", "step": "rendering"}:
        assert time.monotonic() < deadline, "the job never reached rendering"
        time.sleep(0.02)
    GATE.set()
    assert app.settle("evals/demo")["state"]["is"] == "ready"


def test_start_converges_on_what_a_crash_left_behind(repo: Path, app: Client) -> None:
    app.ready()
    app.stop()
    work = repo / "artifacts/review/evals/demo"
    os.rename(work, work.with_name("demo.discard"))
    work.with_name("demo.partial").mkdir()
    (work.with_name("demo.partial") / "changeset.json").write_text("{}")

    reviews = Reviews(render=fake_render, maker_for=maker_for)
    assert sorted(p.name for p in work.parent.iterdir()) == ["demo"]
    assert reviews.row("evals/demo").state.decided == 0
    reviews.close()

    work.with_name("other").mkdir()
    shutil.copytree(work, work.with_name("other.discard"))
    Reviews(render=fake_render, maker_for=maker_for).close()
    assert sorted(p.name for p in work.parent.iterdir()) == ["demo", "other"]


def test_a_second_review_server_is_refused(repo: Path, app: Client, capsys: pytest.CaptureFixture[str]) -> None:
    with pytest.raises(Locked, match=r"^another review server is using artifacts/review/; stop it first$"):
        Reviews(render=fake_render, maker_for=maker_for)

    assert review_main(["serve", "--port", "0"]) == 2
    assert capsys.readouterr().err == "review: another review server is using artifacts/review/; stop it first\n"


def test_a_failed_process_again_keeps_the_review_on_disk(repo: Path, app: Client) -> None:
    app.ready()
    app.decide(c1="keep_old")
    committed = repo / "evals/demo/changeset.json"
    committed.write_text(committed.read_text().replace('"shape": 4', '"shape": 99'))

    app.post("/api/meetings/evals/demo/process", {"again": True})

    assert app.settle("evals/demo")["state"] == {
        "is": "ready",
        "decided": 1,
        "total": 5,
        "applied": False,
        "failed": {
            "message": "The engine refused the ChangeSet: 1 problem.",
            "problems": [
                {
                    "where": "c1 op.shape",
                    "message": "slide 1 (id 256) has no shape 99; shapes with text: 2 'Title 1' ('Market outlook'), "
                    "3 'Content Placeholder 2' ('Demand grows 4% a year Prices…'), 4 'TextBox 3' ('Revenue grew 12% in 2025')",
                }
            ],
        },
    }
    assert app.get("/api/meetings/evals/demo")[1]["review"]["changes"][0]["decision"] == "keep_old"
    assert app.post("/api/meetings/evals/demo/process")[1]["state"]["is"] == "ready"


def test_opening_a_review_whose_source_deck_is_gone_names_the_problem(repo: Path, app: Client) -> None:
    app.ready()
    (repo / "decks/deck.pptx").unlink()

    assert app.get("/api/meetings/evals/demo") == (
        422,
        {"problems": [{"where": "source.path", "message": "cannot read decks/deck.pptx: No such file or directory"}]},
    )
    assert app.decide(c1="keep_old") == (
        422,
        {"problems": [{"where": "source.path", "message": "cannot read decks/deck.pptx: No such file or directory"}]},
    )


def test_a_changeset_the_engine_refuses_fails_with_its_problems(repo: Path, app: Client) -> None:
    write_meeting("bad", [("c1", {**REVENUE, "shape": 99})])

    assert app.ready("evals/bad")["state"] == {
        "is": "failed",
        "message": "The engine refused the ChangeSet: 1 problem.",
        "problems": [
            {
                "where": "c1 op.shape",
                "message": "slide 1 (id 256) has no shape 99; shapes with text: 2 'Title 1' ('Market outlook'), "
                "3 'Content Placeholder 2' ('Demand grows 4% a year Prices…'), 4 'TextBox 3' ('Revenue grew 12% in 2025')",
            }
        ],
    }
    assert not (repo / "artifacts/review/evals/bad").exists()
    assert not (repo / "artifacts/review/evals/bad.partial").exists()


def test_a_missing_source_deck_says_how_to_get_it(repo: Path, app: Client) -> None:
    write_meeting("gone", [("c1", REVENUE)])
    shutil.copyfile(repo / "decks/deck.pptx", repo / "decks/gone.pptx")
    doc = json.loads((repo / "evals/gone/changeset.json").read_text())
    doc["source"]["path"] = "decks/gone.pptx"
    (repo / "evals/gone/changeset.json").write_text(json.dumps(doc))
    (repo / "decks/gone.pptx").unlink()

    state = app.ready("evals/gone")["state"]

    assert state["message"] == (
        "The engine refused the ChangeSet: 1 problem. The source deck is missing; download the corpus decks with "
        "`uv run --project deckcheck python .claude/skills/verify-pptx/scripts/corpus.py`."
    )
    assert state["problems"] == [{"where": "source.path", "message": "cannot read decks/gone.pptx: No such file or directory"}]


def test_a_maker_command_writes_the_changeset_of_a_meeting_without_one(repo: Path, start) -> None:
    (repo / "private/meetings/x").mkdir(parents=True)
    (repo / "private/meetings/x/transcript.md").write_text("# Client sync\n")
    copy = "import shutil,sys;shutil.copyfile(*sys.argv[1:3]);print(sys.argv[3])"
    app = start(functools.partial(maker_for, command=[sys.executable, "-c", copy, "evals/demo/changeset.json", "{dir}/changeset.json", "{meeting}"]))

    row = app.ready("private/x")

    assert row["maker"] == {"label": f"Maker: {sys.executable} -c '{copy}' evals/demo/changeset.json '{{dir}}/changeset.json' '{{meeting}}'", "simulated": False}
    assert row["state"] == {"is": "ready", "decided": 0, "total": 5, "applied": False}
    assert (repo / "artifacts/review/private/x/maker.log").read_text() == "private/meetings/x\n"
    assert app.row("evals/demo")["maker"] == SIMULATED


def test_a_maker_command_that_fails_shows_its_output(repo: Path, start) -> None:
    (repo / "private/meetings/x").mkdir(parents=True)
    (repo / "private/meetings/x/transcript.md").write_text("# Client sync\n")
    app = start(functools.partial(maker_for, command=[sys.executable, "-c", "print('no transcript'); raise SystemExit(3)"]))

    assert app.ready("private/x")["state"] == {"is": "failed", "message": "the maker exited with 3: no transcript", "problems": []}


def test_requests_from_another_site_are_refused(repo: Path, app: Client) -> None:
    assert app.call("GET", "/api/meetings", headers={"Host": "evil.example"}) == (403, {"error": "unknown host"})
    assert app.call("POST", "/api/meetings/evals/demo/process", {}, headers={"Origin": "http://evil.example"}) == (403, {"error": "foreign origin"})
    assert app.call("POST", "/api/meetings/evals/demo/process", {}, headers={"Content-Type": "text/plain"}) == (415, {"error": "send application/json"})
    assert app.call("GET", "/api/meetings", headers={"Sec-Fetch-Site": "cross-site"}) == (403, {"error": "cross-site request"})
    assert app.call("GET", "/api/meetings", headers={"Sec-Fetch-Site": "same-origin"})[0] == 200
    assert app.row("evals/demo")["state"] == {"is": "new"}
    assert app.call("GET", "/api/meetings", headers={"Origin": f"http://localhost:{app.port}", "Host": f"localhost:{app.port}"})[0] == 200
    for path in ("/api/meetings/evals/..%2Fdemo", "/api/meetings/evals/../evals/demo", "/web/../meetings.py", "/web/secret.js", "/web/meetings.py"):
        conn = http.client.HTTPConnection("127.0.0.1", app.port)
        conn.request("GET", path)
        assert conn.getresponse().status == 404, path
        conn.close()


def test_every_response_tells_the_browser_to_keep_it_to_this_page(app: Client) -> None:
    for path in ("/", "/api/meetings", "/api/meetings/evals/nope"):
        conn = http.client.HTTPConnection("127.0.0.1", app.port)
        conn.request("GET", path)
        response = conn.getresponse()
        headers = {name: response.getheader(name) for name in ("X-Frame-Options", "X-Content-Type-Options", "Cross-Origin-Resource-Policy", "Content-Security-Policy")}
        conn.close()

        assert headers == {
            "X-Frame-Options": "DENY",
            "X-Content-Type-Options": "nosniff",
            "Cross-Origin-Resource-Policy": "same-origin",
            "Content-Security-Policy": "default-src 'self'; frame-ancestors 'none'",
        }, path


@pytest.mark.parametrize("length", ["ten", "-1", "1e3"])
def test_a_bad_content_length_is_refused(app: Client, length: str) -> None:
    with socket.create_connection(("127.0.0.1", app.port)) as s:
        s.sendall(
            f"POST /api/meetings/evals/demo/process HTTP/1.1\r\nHost: 127.0.0.1:{app.port}\r\n"
            f"Content-Type: application/json\r\nContent-Length: {length}\r\n\r\n{{}}".encode()
        )
        reply = b"".join(iter(lambda: s.recv(65536), b""))

    assert reply.split(b"\r\n", 1)[0] == b"HTTP/1.0 400 Bad Request"
    assert reply.split(b"\r\n\r\n", 1)[1] == b'{"error": "bad Content-Length"}'
    assert app.row("evals/demo")["state"] == {"is": "new"}


def test_the_log_leaves_out_meeting_names(app: Client, capfd: pytest.CaptureFixture[str]) -> None:
    app.call("POST", "/api/meetings/private/acme-board/process", {}, headers={"Origin": "http://evil.example"})
    app.get("/api/meetings/evals/demo")
    deadline, err = time.monotonic() + 5, ""
    while "GET /api/meetings/<meeting> 409" not in err and time.monotonic() < deadline:
        err += capfd.readouterr().err
        time.sleep(0.02)

    assert err.splitlines()[-2:] == ["POST /api/meetings/<meeting>/process 403", "GET /api/meetings/<meeting> 409"]
    assert "acme" not in err and "demo" not in err


def test_the_app_and_its_files_are_served(app: Client) -> None:
    status, page = app.get("/")
    assert status == 200
    assert b'<script type="module" src="/web/app.js"></script>' in page
    for name in ("app.js", "dom.js", "model.js", "slide.js", "store.js", "style.css"):
        assert app.get(f"/web/{name}") == (200, (WEB / name).read_bytes())


def test_serve_refuses_to_start_outside_the_repo_root(tmp_path: Path, monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]) -> None:
    monkeypatch.chdir(tmp_path)

    assert review_main(["serve", "--port", "0"]) == 2
    assert capsys.readouterr().err == "review: run this from the repo root, the folder that holds deckcheck/ and evals/\n"


def test_no_script_builds_markup_from_a_string() -> None:
    offenders = [p.name for p in sorted(WEB.glob("*.js")) if "innerHTML" in p.read_text() or "insertAdjacentHTML" in p.read_text()]

    assert sorted(p.name for p in WEB.glob("*.js")) == ["app.js", "dom.js", "model.js", "slide.js", "store.js"]
    assert offenders == []


@pytest.mark.skipif(shutil.which("node") is None, reason="node is not installed")
def test_web_model() -> None:
    files = sorted(str(p) for p in (Path(__file__).parent / "web").glob("*.test.mjs"))
    result = subprocess.run(["node", "--test", *files], capture_output=True, text=True)

    assert result.returncode == 0, result.stdout + result.stderr


@pytest.mark.skipif(not (find_soffice() and find_pdftoppm() and find_fc_match()), reason="needs soffice, pdftoppm and fc-match")
def test_process_renders_both_decks_for_real(repo: Path, start) -> None:
    app = start(render=render)
    app.post("/api/meetings/evals/demo/process")
    deadline = time.monotonic() + 120
    while (row := app.row("evals/demo"))["state"]["is"] == "processing":
        assert time.monotonic() < deadline
        time.sleep(0.2)

    assert row["state"]["is"] == "ready", row["state"]
    for side in ("old", "new"):
        status, png = app.get(f"/api/meetings/evals/demo/render/{side}/slide-1.png")
        assert (status, png[:8]) == (200, b"\x89PNG\r\n\x1a\n")
