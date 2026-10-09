"""The maker seam. A maker reads a meeting and writes a ChangeSet; the engine checks it next.

Piece 3 plugs in here without touching the server: `review serve --maker 'claude -p ... {meeting} ... {out}'`
runs that command for every meeting. Without `--maker`, a golden scenario replays its committed changeset.json.
"""

from __future__ import annotations

import shlex
import shutil
import subprocess
from collections.abc import Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import TYPE_CHECKING, ClassVar, Protocol

if TYPE_CHECKING:
    from deckcheck.review.meetings import Meeting

TIMEOUT = 30 * 60
TAIL = 2000


class MakerFailed(Exception):
    pass


class Maker(Protocol):
    simulated: bool
    label: str

    def make(self, meeting: Meeting, out: Path) -> None: ...


@dataclass(frozen=True)
class Golden:
    simulated: ClassVar[bool] = True
    label: ClassVar[str] = "Simulated maker: replays the committed changeset.json"

    def make(self, meeting: Meeting, out: Path) -> None:
        shutil.copyfile(meeting.dir / "changeset.json", out)


@dataclass(frozen=True)
class Command:
    argv: tuple[str, ...]
    simulated: ClassVar[bool] = False

    @property
    def label(self) -> str:
        return f"Maker: {shlex.join(self.argv)}"

    def make(self, meeting: Meeting, out: Path) -> None:
        argv = [a.replace("{meeting}", str(meeting.dir)).replace("{out}", str(out)) for a in self.argv]
        log = out.with_name("maker.log")
        with log.open("wb") as f:
            try:
                code = subprocess.run(argv, stdout=f, stderr=subprocess.STDOUT, timeout=TIMEOUT, check=False).returncode
            except subprocess.TimeoutExpired:
                raise MakerFailed(f"the maker ran longer than {TIMEOUT // 60} minutes and was stopped") from None
            except OSError as e:
                raise MakerFailed(f"cannot run the maker {argv[0]!r}: {e.strerror}") from None
        if code != 0:
            tail = log.read_bytes()[-TAIL:].decode("utf-8", errors="replace").strip()
            raise MakerFailed(f"the maker exited with {code}" + (f": {tail}" if tail else ""))


def maker_for(meeting: Meeting, command: Sequence[str] | None = None) -> Maker | None:
    if command:
        return Command(tuple(command))
    if (meeting.dir / "changeset.json").is_file():
        return Golden()
    return None
