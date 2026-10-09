from __future__ import annotations

import json
import re
import sys
from collections.abc import Callable
from dataclasses import asdict, dataclass
from email.message import Message
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlsplit

from pydantic import BaseModel, ConfigDict, ValidationError

from deckcheck.changeset import Invalid, Undecided
from deckcheck.changeset.model import Decision
from deckcheck.review.meetings import Conflict, Failed, Final, New, Processing, Ready, Reviews, Row, State, Unknown, View

WEB = Path(__file__).parent / "web"
TEXT = "; charset=utf-8"
JSON = "application/json" + TEXT
STATIC = {
    "index.html": "text/html" + TEXT,
    "style.css": "text/css" + TEXT,
    "app.js": "text/javascript" + TEXT,
    "dom.js": "text/javascript" + TEXT,
    "model.js": "text/javascript" + TEXT,
    "slide.js": "text/javascript" + TEXT,
    "store.js": "text/javascript" + TEXT,
}
PPTX = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
MAX_BODY = 1 << 20
MEETING = r"/api/meetings/(?P<mid>(?:evals|private)/[A-Za-z0-9][A-Za-z0-9_.-]*)"
NOT_FOUND = 404, {"error": "not found"}
HEADERS = {
    "X-Frame-Options": "DENY",
    "X-Content-Type-Options": "nosniff",
    "Cross-Origin-Resource-Policy": "same-origin",
    "Content-Security-Policy": "default-src 'self'; frame-ancestors 'none'",
}


class Body(BaseModel):
    model_config = ConfigDict(extra="forbid", strict=True)


class ProcessBody(Body):
    again: bool = False


class DecisionsBody(Body):
    decisions: dict[str, Decision]


@dataclass(frozen=True)
class File:
    path: Path
    type: str
    download: str | None = None
    cache: str = "no-store"


@dataclass(frozen=True)
class Blob:
    data: bytes
    type: str = JSON
    cache: str = "no-store"
    download: str | None = None


Reply = tuple[int, object]


class App:
    def __init__(self, reviews: Reviews) -> None:
        self.reviews = reviews
        self.routes: list[tuple[str, re.Pattern[str], Callable[..., Reply]]] = [
            ("GET", re.compile(r"/"), lambda body: (200, File(WEB / "index.html", STATIC["index.html"]))),
            ("GET", re.compile(r"/web/(?P<file>[a-z]+\.(?:html|css|js))"), self.static),
            ("GET", re.compile(r"/api/meetings"), lambda body: (200, {"meetings": [row_json(r) for r in reviews.meetings()]})),
            ("POST", re.compile(MEETING + r"/process"), self.process),
            ("GET", re.compile(MEETING), self.meeting),
            ("POST", re.compile(MEETING + r"/decisions"), self.decide),
            ("POST", re.compile(MEETING + r"/apply"), self.apply),
            ("GET", re.compile(MEETING + r"/render/(?P<side>old|new)/slide-(?P<n>[1-9][0-9]{0,3})\.png"), self.png),
            ("GET", re.compile(MEETING + r"/final\.pptx"), self.final),
        ]

    def static(self, body: bytes, file: str) -> Reply:
        if file not in STATIC:
            raise Unknown(f"no file {file}")
        return 200, File(WEB / file, STATIC[file])

    def process(self, body: bytes, mid: str) -> Reply:
        return 202, row_json(self.reviews.process(mid, ProcessBody.model_validate_json(body).again))

    def meeting(self, body: bytes, mid: str) -> Reply:
        row = self.reviews.row(mid)
        if not isinstance(row.state, Ready):
            return 409, row_json(row)
        return 200, view_json(self.reviews.view(mid))

    def decide(self, body: bytes, mid: str) -> Reply:
        decisions, final = self.reviews.decide(mid, DecisionsBody.model_validate_json(body).decisions)
        wire = {cid: d if isinstance(d, str) else d.model_dump() for cid, d in decisions.items()}
        return 200, {"decisions": wire, "final": final_json(final, mid)}

    def apply(self, body: bytes, mid: str) -> Reply:
        Body.model_validate_json(body)
        return 200, final_json(self.reviews.apply(mid), mid)

    def png(self, body: bytes, mid: str, side: str, n: str) -> Reply:
        return 200, File(self.reviews.slide_png(mid, side, int(n)), "image/png", cache="private, max-age=86400")

    def final(self, body: bytes, mid: str) -> Reply:
        name = mid.split("/", 1)[1]
        return 200, File(self.reviews.final_pptx(mid), PPTX, download=f"{name}-final.pptx")

    def handle(self, method: str, path: str, body: bytes) -> Reply:
        for verb, pattern, handler in self.routes:
            if verb == method and (m := pattern.fullmatch(path)):
                try:
                    status, payload = handler(body, **m.groupdict())
                except Exception as e:
                    for kind, reply in ERRORS.items():
                        if isinstance(e, kind):
                            return reply(e)
                    print(f"review: {type(e).__name__} on {method} {redacted(path)}", file=sys.stderr)
                    return 500, {"error": f"internal error: {e}"}
                if not isinstance(payload, File):
                    return status, payload
                try:
                    return status, Blob(payload.path.read_bytes(), payload.type, payload.cache, payload.download)
                except OSError:
                    return NOT_FOUND
        return NOT_FOUND


def redacted(path: str) -> str:
    return re.sub(r"^/api/meetings/[^/]+/[^/]+", "/api/meetings/<meeting>", path)


ERRORS: dict[type[Exception], Callable[..., Reply]] = {
    Unknown: lambda e: (404, {"error": str(e)}),
    Conflict: lambda e: (409, {"error": str(e)}),
    Undecided: lambda e: (409, {"error": "decide every change first", "pending": list(e.pending)}),
    Invalid: lambda e: (422, {"problems": [asdict(p) for p in e.problems]}),
    ValidationError: lambda e: (400, {"error": "bad request body", "problems": [err["msg"] for err in e.errors()]}),
}


def row_json(row: Row) -> dict:
    m = row.meeting
    return {
        "id": m.id,
        "origin": m.origin,
        "name": m.name,
        "title": m.title,
        "date": m.date,
        "maker": {"label": row.maker.label, "simulated": row.maker.simulated} if row.maker else None,
        "state": state_json(row.state),
        "fonts": [{"font": font, "family": family} for font, family in row.fonts],
    }


def state_json(state: State) -> dict:
    match state:
        case New():
            return {"is": "new"}
        case Processing(step):
            return {"is": "processing", "step": step}
        case Failed(message, problems):
            return {"is": "failed", "message": message, "problems": [asdict(p) for p in problems]}
        case Ready(decided, total, applied):
            return {"is": "ready", "decided": decided, "total": total, "applied": applied}


def final_json(final: Final | None, mid: str) -> dict | None:
    return {**asdict(final), "download": f"/api/meetings/{mid}/final.pptx"} if final else None


def view_json(view: View) -> dict:
    mid = view.row.meeting.id
    stamp = view.review.executed.sha256[:12]
    w, h = view.slide
    return {
        **row_json(view.row),
        "review": view.review.model_dump(mode="json"),
        "slide": {"w": w, "h": h},
        "images": {side: f"/api/meetings/{mid}/render/{side}/slide-{{n}}.png?v={stamp}" for side in ("old", "new")},
        "final": final_json(view.final, mid),
    }


def serve(reviews: Reviews, port: int = 8765) -> ThreadingHTTPServer:
    app = App(reviews)

    class Handler(BaseHTTPRequestHandler):
        def do_GET(self) -> None:
            self._serve("GET")

        def do_POST(self) -> None:
            self._serve("POST")

        def log_message(self, format: str, *args: object) -> None:
            pass

        def _serve(self, method: str) -> None:
            path = urlsplit(self.path).path
            length = _length(self.headers.get("Content-Length"))
            if refused := _guard(method, self.headers, server.server_address[1]):
                reply = refused
            elif length is None:
                reply = 400, {"error": "bad Content-Length"}
            elif length > MAX_BODY:
                reply = 413, {"error": "request body too large"}
            else:
                reply = app.handle(method, path, self.rfile.read(length) if method == "POST" else b"")
            self._send(*reply)
            print(f"{method} {redacted(path)} {reply[0]}", file=sys.stderr, flush=True)

        def _send(self, status: int, payload: object) -> None:
            blob = payload if isinstance(payload, Blob) else Blob(json.dumps(payload, ensure_ascii=False).encode())
            self.send_response(status)
            for name, value in HEADERS.items():
                self.send_header(name, value)
            self.send_header("Content-Type", blob.type)
            self.send_header("Content-Length", str(len(blob.data)))
            self.send_header("Cache-Control", blob.cache)
            if blob.download:
                self.send_header("Content-Disposition", f'attachment; filename="{blob.download}"')
            self.end_headers()
            self.wfile.write(blob.data)

    server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    server.daemon_threads = True
    return server


def _length(header: str | None) -> int | None:
    try:
        length = int(header or 0)
    except ValueError:
        return None
    return length if length >= 0 else None


def _guard(method: str, headers: Message, port: int) -> Reply | None:
    """Host stops DNS rebinding; Origin, Sec-Fetch-Site and the JSON content type stop another site's page reaching us."""
    hosts = {f"127.0.0.1:{port}", f"localhost:{port}"}
    if headers.get("Host") not in hosts:
        return 403, {"error": "unknown host"}
    if headers.get("Sec-Fetch-Site") == "cross-site":
        return 403, {"error": "cross-site request"}
    origin = headers.get("Origin")
    if origin is not None and origin not in {f"http://{h}" for h in hosts}:
        return 403, {"error": "foreign origin"}
    if method == "POST" and headers.get_content_type() != "application/json":
        return 415, {"error": "send application/json"}
    return None
