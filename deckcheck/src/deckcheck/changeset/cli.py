from __future__ import annotations

import argparse
import hashlib
import os
import sys
from pathlib import Path

from deckcheck.changeset.engine import Checked, Invalid, Item, Undecided, apply, execute, load, review
from deckcheck.changeset.model import ReplaceText
from deckcheck.cli import write_atomic
from deckcheck.fix import plural
from deckcheck.model import DeckError

OK, PROBLEMS, USAGE = 0, 1, 2


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="changeset", description="Check a ChangeSet against its deck, run it, and apply review decisions.")
    sub = parser.add_subparsers(dest="command", required=True)
    validate = sub.add_parser("validate", help="check every change against the source deck")
    validate.add_argument("changeset", type=Path)
    run = sub.add_parser("execute", help="write the deck with every change as the maker wrote it")
    run.add_argument("changeset", type=Path)
    run.add_argument("--out", type=Path, required=True, help="the executed deck, a new .pptx path")
    run.add_argument("--review", type=Path, help="write the review view, a JSON file, here")
    final = sub.add_parser("apply", help="write the final deck from the source and the review decisions")
    final.add_argument("changeset", type=Path)
    final.add_argument("--out", type=Path, required=True, help="the final deck, a new .pptx path")
    args = parser.parse_args(argv)
    try:
        checked = load(args.changeset)
        match args.command:
            case "validate":
                return cmd_validate(checked)
            case "execute":
                return cmd_execute(checked, args.out, args.review)
            case _:
                return cmd_apply(checked, args.out)
    except Invalid as e:
        print(f"INVALID {args.changeset}: {plural(len(e.problems), 'problem', 'problems')}")
        for p in e.problems:
            print(f"  {p.where}: {p.message}")
        return PROBLEMS
    except Undecided as e:
        print(f"PENDING {args.changeset}: {plural(len(e.pending), 'decision', 'decisions')} pending ({', '.join(e.pending)}); nothing written")
        return PROBLEMS
    except (DeckError, OSError) as e:
        print(f"error: {e}", file=sys.stderr)
        return USAGE


def cmd_validate(checked: Checked) -> int:
    cs = checked.changeset
    keys = sorted({i.source_index for i in checked.items if i.source_index is not None})
    added = [str(i.slide) for i in checked.items if i.source_index is None and i.slide == i.change.id]
    where = [*map(str, keys), *(f"new {a}" for a in added)]
    print(
        f"VALID {checked.path}: {_counts(checked)} on {'slide' if len(where) == 1 else 'slides'} {', '.join(where)}; "
        f"{plural(len(cs.asks), 'ask', 'asks')}, {plural(len(cs.flags), 'flag', 'flags')}, {len(cs.held)} held"
    )
    for item in checked.items:
        print(f"  {format_item(item)}")
    return OK


def _counts(checked: Checked) -> str:
    structural = sum(i.change.op.structural for i in checked.items)
    return f"{plural(len(checked.items), 'change', 'changes')} ({len(checked.items) - structural} text-only, {structural} structural)"


def format_item(item: Item) -> str:
    op = item.change.op
    slide = f"slide {item.source_index}" if item.source_index is not None else f"new slide {item.slide}"
    shape = f" shape {item.shape[0]} {item.shape[1]!r}" if item.shape else ""
    before, after = (op.old, op.new) if isinstance(op, ReplaceText) else (item.before, item.after)
    kind = "structural" if op.structural else "text-only"
    return f"{item.change.id} {op.kind} {slide}{shape}: {_show(before)} -> {_show(after)} ({kind})"


def _show(value: str | None) -> str:
    if value is None:
        return "nothing"
    return repr(value if len(value) <= 80 else value[:79] + "…")


def _same(a: Path, b: Path) -> bool:
    if a.exists() and b.exists():
        return os.path.samefile(a, b)
    return a.resolve() == b.resolve()


def _refuse_overwrites(checked: Checked, outputs: dict[str, Path]) -> None:
    inputs = {"the source deck": Path(checked.changeset.source.path), "the ChangeSet": checked.path}
    for flag, path in outputs.items():
        if path.is_dir():
            raise DeckError(f"{flag} {path} is a directory; pass the path of a new file")
        for what, source in inputs.items():
            if _same(path, source):
                raise DeckError(f"{flag} {path} is {what}; the engine never writes over it")
    if "--review" in outputs and _same(outputs["--out"], outputs["--review"]):
        raise DeckError(f"--review {outputs['--review']} is also --out; give each its own path")


def cmd_execute(checked: Checked, out: Path, review_path: Path | None) -> int:
    _refuse_overwrites(checked, {"--out": out} | ({"--review": review_path} if review_path else {}))
    data = execute(checked)
    write_atomic(out, data)
    print(f"EXECUTED {checked.path} -> {out}: {_counts(checked)} written, sha256 {hashlib.sha256(data).hexdigest()}")
    for item in checked.items:
        for note in item.notes:
            print(f"  {item.change.id}: {note}")
    if review_path:
        doc = review(checked, out, data)
        write_atomic(review_path, (doc.model_dump_json(indent=2) + "\n").encode())
        print(f"review -> {review_path}")
    return OK


def cmd_apply(checked: Checked, out: Path) -> int:
    _refuse_overwrites(checked, {"--out": out})
    result = apply(checked)
    write_atomic(out, result.data)
    print(
        f"APPLIED {checked.path} -> {out}: {len(result.kept_new)} kept new, {len(result.edited)} edited, "
        f"{len(result.kept_old)} kept old, {len(result.dropped)} dropped, sha256 {hashlib.sha256(result.data).hexdigest()}"
    )
    for cid, add in result.dropped:
        print(f"  dropped {cid}: it fills the slide {add} adds, which was kept old")
    return OK
