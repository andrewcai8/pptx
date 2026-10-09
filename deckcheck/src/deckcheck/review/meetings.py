from __future__ import annotations

import copy
import fcntl
import hashlib
import io
import itertools
import json
import os
import re
import shutil
import threading
from collections.abc import Callable, Mapping, Sequence
from concurrent.futures import ThreadPoolExecutor
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Literal

from pptx import Presentation

from deckcheck.changeset import Invalid, Problem, Review, apply, check, execute, load, parse, review
from deckcheck.changeset.model import Decision, Edited
from deckcheck.cli import write_atomic
from deckcheck.fix import plural
from deckcheck.review.maker import Maker

Origin = Literal["evals", "private"]
Step = Literal["making", "executing", "rendering"]
Side = Literal["old", "new"]
Render = Callable[[Path, Path], object]

ROOTS: dict[Origin, Path] = {"evals": Path("evals"), "private": Path("private/meetings")}
WORK = Path("artifacts/review")
NAME = re.compile(r"[A-Za-z0-9][A-Za-z0-9_.-]*(?<!\.partial|\.discard)")
DATE = re.compile(r"^Date:\s*(\d{4}-\d{2}-\d{2})")
CORPUS = "download the corpus decks with `uv run --project deckcheck python .claude/skills/verify-pptx/scripts/corpus.py`"


@dataclass(frozen=True)
class Meeting:
    origin: Origin
    name: str
    dir: Path
    title: str
    date: str | None

    @property
    def id(self) -> str:
        return f"{self.origin}/{self.name}"


@dataclass(frozen=True)
class New:
    pass


@dataclass(frozen=True)
class Processing:
    step: Step


@dataclass(frozen=True)
class Failed:
    message: str
    problems: tuple[Problem, ...] = ()


@dataclass(frozen=True)
class Ready:
    decided: int
    total: int
    applied: bool
    failed: Failed | None = None


State = New | Processing | Failed | Ready


@dataclass(frozen=True)
class Final:
    path: str
    sha256: str
    changeset_sha256: str
    kept_new: tuple[str, ...]
    edited: tuple[str, ...]
    kept_old: tuple[str, ...]
    dropped: tuple[tuple[str, str], ...]


@dataclass(frozen=True)
class Row:
    meeting: Meeting
    maker: Maker | None
    state: State
    fonts: tuple[tuple[str, str], ...]


@dataclass(frozen=True)
class View:
    row: Row
    review: Review
    slide: tuple[int, int]
    stamp: str
    final: Final | None


class Unknown(Exception):
    pass


class Conflict(Exception):
    pass


class Locked(Exception):
    pass


def discover() -> list[Meeting]:
    return [_meeting(origin, d) for origin, root in ROOTS.items() if root.is_dir() for d in sorted(root.iterdir()) if _holds_meeting(d)]


def find(mid: str) -> Meeting:
    origin, _, name = mid.partition("/")
    if origin in ROOTS and NAME.fullmatch(name) and _holds_meeting(d := ROOTS[origin] / name):
        return _meeting(origin, d)
    raise Unknown(f"no meeting {mid}")


def _holds_meeting(d: Path) -> bool:
    return bool(NAME.fullmatch(d.name)) and ((d / "transcript.md").is_file() or (d / "changeset.json").is_file())


def _meeting(origin: Origin, d: Path) -> Meeting:
    title = date = None
    transcript = d / "transcript.md"
    if transcript.is_file():
        with transcript.open(encoding="utf-8", errors="replace") as f:
            for line in itertools.islice(f, 40):
                if title is None and line.startswith("# "):
                    title = line[2:].strip() or None
                if date is None and (match := DATE.match(line)):
                    date = match[1]
    if title is None or date is None:
        meeting = _changeset_meeting(d / "changeset.json")
        title = title or meeting.get("title")
        date = date or meeting.get("date")
    return Meeting(origin, d.name, d, title or d.name, date)


def _changeset_meeting(path: Path) -> dict[str, str]:
    try:
        meeting = json.loads(path.read_text(encoding="utf-8-sig")).get("meeting")
    except (OSError, ValueError, AttributeError):
        return {}
    return {k: v for k, v in meeting.items() if isinstance(v, str)} if isinstance(meeting, dict) else {}


def cascade(changes: Sequence[Mapping], batch: Mapping[str, Decision]) -> dict[str, Decision]:
    """Dropping an added slide sets its pending fills to keep_old; restoring it reopens its keep_old fills.

    Any other fill decision stands, so apply reports a kept fill of a dropped slide under dropped."""
    now = {c["id"]: c.get("decision", "pending") for c in changes}
    out = dict(batch)
    for c in changes:
        op = c["op"]
        if op["kind"] != "fill_placeholder" or c["id"] in batch or op["slide"] not in batch:
            continue
        add = op["slide"]
        if batch[add] == "keep_old" and now[c["id"]] == "pending":
            out[c["id"]] = "keep_old"
        elif now[add] == "keep_old" and batch[add] != "keep_old" and now[c["id"]] == "keep_old":
            out[c["id"]] = "pending"
    return out


@dataclass(frozen=True)
class _Workdir:
    dir: Path

    @classmethod
    def of(cls, meeting: Meeting) -> _Workdir:
        return cls(WORK / meeting.origin / meeting.name)

    @property
    def partial(self) -> Path:
        return self.dir.with_name(self.dir.name + ".partial")

    @property
    def discard(self) -> Path:
        return self.dir.with_name(self.dir.name + ".discard")

    @property
    def changeset(self) -> Path:
        return self.dir / "changeset.json"

    @property
    def executed(self) -> Path:
        return self.dir / "executed.pptx"

    @property
    def final_pptx(self) -> Path:
        return self.dir / "final.pptx"

    @property
    def applied(self) -> Path:
        return self.dir / "applied.json"

    @property
    def renders(self) -> Path:
        return self.dir / "render"

    def render(self, side: Side) -> Path:
        return self.renders / side

    def final(self) -> Final | None:
        try:
            doc = json.loads(self.applied.read_text())
            current = _sha(self.changeset.read_bytes())
        except OSError:
            return None
        if doc["changeset_sha256"] != current or not self.final_pptx.is_file():
            return None
        lists = {k: tuple(doc[k]) for k in ("kept_new", "edited", "kept_old")}
        return Final(doc["path"], doc["sha256"], doc["changeset_sha256"], **lists, dropped=tuple(map(tuple, doc["dropped"])))

    def fonts(self) -> tuple[tuple[str, str], ...]:
        found = set()
        for side in ("old", "new"):
            try:
                report = json.loads((self.render(side) / "fonts.json").read_text())
            except OSError:
                continue
            found |= {(font, m["family"]) for font, m in report.items() if m["substituted"]}
        return tuple(sorted(found))


@dataclass
class _Job:
    step: Step = "making"


class Reviews:

    def __init__(self, *, render: Render, maker_for: Callable[[Meeting], Maker | None]) -> None:
        self._render = render
        self._maker_for = maker_for
        self._lock = threading.Lock()
        self._jobs: dict[str, _Job] = {}
        self._failed: dict[str, Failed] = {}
        self._locks: dict[str, threading.Lock] = {}
        self._fd = _hold(WORK / ".lock")
        _converge()

    def close(self) -> None:
        if self._fd is not None:
            os.close(self._fd)
            self._fd = None

    def meetings(self) -> list[Row]:
        return [self._row(m) for m in discover()]

    def row(self, mid: str) -> Row:
        return self._row(find(mid))

    def process(self, mid: str, again: bool = False) -> Row:
        meeting = find(mid)
        maker = self._maker_for(meeting)
        if maker is None:
            raise Conflict("this meeting has no changeset.json; start the server with --maker to process it")
        with self._lock:
            idle = mid not in self._jobs
            start = idle and (again or not _Workdir.of(meeting).dir.is_dir())
            if start:
                self._failed.pop(mid, None)
                job = self._jobs[mid] = _Job()
        if start:
            threading.Thread(target=self._run, args=(meeting, maker, job), daemon=True).start()
        return self._row(meeting)

    def view(self, mid: str) -> View | Row:
        meeting = find(mid)
        row = self._row(meeting)
        if not isinstance(row.state, Ready):
            return row
        work = _Workdir.of(meeting)
        with self._meeting_lock(mid):
            self._idle(mid)
            checked = load(work.changeset)
            doc = review(checked, work.executed, work.executed.read_bytes())
            final = work.final()
            stamp = _sha(f"{doc.executed.sha256} {work.renders.stat().st_mtime_ns}".encode())[:12]
        deck = Presentation(io.BytesIO(checked.source))
        return View(row, doc, (deck.slide_width, deck.slide_height), stamp, final)

    def decide(self, mid: str, decisions: Mapping[str, Decision]) -> tuple[dict[str, Decision], Final | None]:
        work = self._ready(find(mid))
        with self._meeting_lock(mid):
            self._idle(mid)
            raw = json.loads(work.changeset.read_text(encoding="utf-8-sig"))
            changes = raw["changes"]
            known = {c["id"] for c in changes}
            if unknown := [cid for cid in decisions if cid not in known]:
                raise Invalid([Problem(cid, "no such change") for cid in unknown])
            before = copy.deepcopy(changes)
            updates = cascade(changes, decisions)
            for c in changes:
                if c["id"] in updates:
                    _set_decision(c, updates[c["id"]])
            if changes != before:
                text = json.dumps(raw, indent=2, ensure_ascii=False) + "\n"
                _check(text, work.changeset)
                write_atomic(work.changeset, text.encode())
            return {c["id"]: _decision(c) for c in changes}, work.final()

    def apply(self, mid: str) -> Final:
        work = self._ready(find(mid))
        with self._meeting_lock(mid):
            self._idle(mid)
            loaded = work.changeset.read_bytes()
            result = apply(load(work.changeset))
            work.applied.unlink(missing_ok=True)
            write_atomic(work.final_pptx, result.data)
            final = Final(
                path=str(work.final_pptx),
                sha256=_sha(result.data),
                changeset_sha256=_sha(loaded),
                kept_new=result.kept_new,
                edited=result.edited,
                kept_old=result.kept_old,
                dropped=result.dropped,
            )
            write_atomic(work.applied, (json.dumps(asdict(final), indent=2) + "\n").encode())
            return final

    def slide_png(self, mid: str, side: Side, n: int) -> Path:
        return _Workdir.of(find(mid)).render(side) / f"slide-{n}.png"

    def final_pptx(self, mid: str) -> Path:
        work = _Workdir.of(find(mid))
        if work.final() is None:
            raise Unknown("no final deck for the current decisions; apply them first")
        return work.final_pptx

    def _row(self, meeting: Meeting) -> Row:
        work = _Workdir.of(meeting)
        return Row(meeting, self._maker_for(meeting), self._state(meeting, work), work.fonts())

    def _state(self, meeting: Meeting, work: _Workdir) -> State:
        with self._lock:
            job, failed = self._jobs.get(meeting.id), self._failed.get(meeting.id)
        if job:
            return Processing(job.step)
        try:
            changes = json.loads(work.changeset.read_text(encoding="utf-8-sig"))["changes"]
        except FileNotFoundError:
            return failed or New()
        decided = sum(c.get("decision", "pending") != "pending" for c in changes)
        return Ready(decided, len(changes), work.final() is not None, failed)

    def _ready(self, meeting: Meeting) -> _Workdir:
        work = _Workdir.of(meeting)
        if not work.changeset.is_file():
            raise Conflict("this meeting has not been processed yet")
        return work

    def _idle(self, mid: str) -> None:
        with self._lock:
            if mid in self._jobs:
                raise Conflict("this meeting is being processed; wait until it is ready")

    def _meeting_lock(self, mid: str) -> threading.Lock:
        with self._lock:
            return self._locks.setdefault(mid, threading.Lock())

    def _run(self, meeting: Meeting, maker: Maker, job: _Job) -> None:
        work = _Workdir.of(meeting)
        build = _Workdir(work.partial)
        failed = None
        try:
            shutil.rmtree(build.dir, ignore_errors=True)
            build.dir.mkdir(parents=True)
            maker.make(meeting, build.changeset)
            job.step = "executing"
            checked = load(build.changeset)
            write_atomic(build.executed, execute(checked))
            job.step = "rendering"
            sources = {"old": Path(checked.changeset.source.path), "new": build.executed}
            with ThreadPoolExecutor(len(sources)) as pool:
                for done in [pool.submit(self._render, deck, build.render(side)) for side, deck in sources.items()]:
                    done.result()
            with self._meeting_lock(meeting.id):
                if work.dir.exists():
                    os.rename(work.dir, work.discard)
                os.rename(build.dir, work.dir)
                shutil.rmtree(work.discard, ignore_errors=True)
        except Invalid as e:
            failed = _refused(meeting, e)
        except Exception as e:
            failed = Failed(str(e) or type(e).__name__)
        finally:
            shutil.rmtree(build.dir, ignore_errors=True)
            with self._lock:
                del self._jobs[meeting.id]
                if failed:
                    self._failed[meeting.id] = failed


def _hold(path: Path) -> int:
    path.parent.mkdir(parents=True, exist_ok=True)
    fd = os.open(path, os.O_RDWR | os.O_CREAT, 0o666)
    try:
        fcntl.flock(fd, fcntl.LOCK_EX | fcntl.LOCK_NB)
    except BlockingIOError:
        os.close(fd)
        raise Locked(f"another review server is using {WORK}/; stop it first") from None
    return fd


def _converge() -> None:
    for origin in ROOTS:
        root = WORK / origin
        if not root.is_dir():
            continue
        for p in root.iterdir():
            if p.name.endswith(".partial"):
                shutil.rmtree(p)
            elif p.name.endswith(".discard"):
                live = p.with_name(p.name.removesuffix(".discard"))
                if live.exists():
                    shutil.rmtree(p)
                else:
                    os.rename(p, live)


def _refused(meeting: Meeting, e: Invalid) -> Failed:
    message = f"The engine refused the ChangeSet: {plural(len(e.problems), 'problem', 'problems')}."
    if any(p.where == "source.path" for p in e.problems):
        message += f" The source deck is missing; {CORPUS}."
    return Failed(message, e.problems)


def _decision(change: Mapping) -> Decision:
    d = change.get("decision", "pending")
    return Edited(edited=d["edited"]) if isinstance(d, dict) else d


def _set_decision(change: dict, decision: Decision) -> None:
    if decision == "pending":
        change.pop("decision", None)
    elif isinstance(decision, Edited):
        change["decision"] = {"edited": decision.edited}
    else:
        change["decision"] = decision


def _check(text: str, path: Path) -> None:
    cs = parse(text)
    try:
        source = Path(cs.source.path).read_bytes()
    except OSError as e:
        raise Invalid([Problem("source.path", f"cannot read {cs.source.path}: {e.strerror}")]) from e
    check(cs, source, path)


def _sha(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()
