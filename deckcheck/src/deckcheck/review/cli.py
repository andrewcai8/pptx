from __future__ import annotations

import argparse
import functools
import shlex
import sys
from pathlib import Path

from deckcheck.render import render
from deckcheck.review.maker import maker_for
from deckcheck.review.meetings import Reviews
from deckcheck.review.server import serve

USAGE = 2


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="review", description="Serve the review app, where a consultant decides each change a maker made.")
    sub = parser.add_subparsers(dest="command", required=True)
    run = sub.add_parser("serve", help="serve the app on 127.0.0.1")
    run.add_argument("--port", type=int, default=8765, help="0 picks a free port")
    run.add_argument("--maker", help="a command that writes a ChangeSet, with {meeting} and {out} replaced for each meeting")
    args = parser.parse_args(argv)
    if not (Path("deckcheck/pyproject.toml").is_file() and Path("evals").is_dir()):
        print("review: run this from the repo root, the folder that holds deckcheck/ and evals/", file=sys.stderr)
        return USAGE
    command = shlex.split(args.maker) if args.maker is not None else None
    if command == []:
        print("review: --maker is empty", file=sys.stderr)
        return USAGE
    try:
        server = serve(Reviews(render=render, maker_for=functools.partial(maker_for, command=command)), args.port)
    except OSError as e:
        print(f"review: cannot listen on 127.0.0.1:{args.port}: {e.strerror}", file=sys.stderr)
        return USAGE
    print(f"review: http://127.0.0.1:{server.server_address[1]}", flush=True)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
    return 0
