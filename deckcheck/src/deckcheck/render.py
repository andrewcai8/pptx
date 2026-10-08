from __future__ import annotations

import json
import re
import shutil
import subprocess
import tempfile
from dataclasses import asdict, dataclass
from pathlib import Path

from deckcheck.model import Deck, load_deck

MAC_SOFFICE = Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")
# Metric-compatible clones share glyph widths with the requested font, so line breaks match.
METRIC_CLONES = {
    "calibri": {"carlito"},
    "cambria": {"caladea"},
    "arial": {"liberation sans"},
    "helvetica": {"liberation sans", "arial"},
    "times new roman": {"liberation serif"},
    "courier new": {"liberation mono"},
}


class ToolMissing(Exception):
    pass


class RenderError(Exception):
    pass


@dataclass(frozen=True)
class FontMatch:
    family: str
    substituted: bool


def find_soffice() -> str | None:
    return shutil.which("soffice") or (str(MAC_SOFFICE) if MAC_SOFFICE.exists() else None)


def find_pdftoppm() -> str | None:
    return shutil.which("pdftoppm")


def find_fc_match() -> str | None:
    return shutil.which("fc-match")


def require_tools() -> tuple[str, str]:
    if find_fc_match() is None:
        raise ToolMissing("fc-match not found; install fontconfig with `brew install fontconfig`")
    soffice = find_soffice()
    if soffice is None:
        raise ToolMissing("soffice not found; install LibreOffice with `brew install --cask libreoffice`")
    pdftoppm = find_pdftoppm()
    if pdftoppm is None:
        raise ToolMissing("pdftoppm not found; install poppler with `brew install poppler`")
    return soffice, pdftoppm


def font_report(deck: Deck) -> dict[str, FontMatch]:
    typefaces = {r.font for s in deck.slides for sh in s.shapes for p in sh.paragraphs for r in p.runs} - {""}
    report = {}
    for typeface in sorted(typefaces):
        result = subprocess.run(["fc-match", "-f", "%{family}", typeface], capture_output=True, text=True, check=False)
        families = [f.strip() for f in result.stdout.split(",")]
        matched = {f.lower() for f in families} & ({typeface.lower()} | METRIC_CLONES.get(typeface.lower(), set()))
        report[typeface] = FontMatch(family=families[0], substituted=not matched)
    return report


def render(deck: Path, out: Path) -> tuple[list[Path], dict[str, FontMatch]]:
    soffice, pdftoppm = require_tools()
    fonts = font_report(load_deck(deck))
    out.mkdir(parents=True, exist_ok=True)
    (out / "fonts.json").write_text(json.dumps({k: asdict(v) for k, v in fonts.items()}, indent=2) + "\n")
    # A private profile keeps headless soffice from silently no-oping when a desktop LibreOffice is open.
    with tempfile.TemporaryDirectory() as profile:
        _run([
            soffice,
            f"-env:UserInstallation={Path(profile).as_uri()}",
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            str(out),
            str(deck),
        ])
    pdf = out / f"{deck.stem}.pdf"
    if not pdf.exists():
        raise RenderError(f"soffice produced no PDF for {deck}")
    for stale in out.glob("slide-*.png"):
        stale.unlink()
    _run([pdftoppm, "-png", "-r", "80", str(pdf), str(out / "slide")])
    # pdftoppm zero-pads page numbers to the page count's width; normalize to slide-1.png.
    pngs = []
    for png in out.glob("slide-*.png"):
        n = int(re.fullmatch(r"slide-(\d+)\.png", png.name).group(1))
        pngs.append(png.rename(out / f"slide-{n}.png"))
    return sorted(pngs, key=lambda p: int(p.stem.split("-")[1])), fonts


def _run(cmd: list[str]) -> None:
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        raise RenderError(f"{Path(cmd[0]).name} failed ({result.returncode}): {result.stderr.strip()}")
