from __future__ import annotations

import re
import shutil
import subprocess
import tempfile
from pathlib import Path

MAC_SOFFICE = Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")


class ToolMissing(Exception):
    pass


class RenderError(Exception):
    pass


def find_soffice() -> str | None:
    return shutil.which("soffice") or (str(MAC_SOFFICE) if MAC_SOFFICE.exists() else None)


def find_pdftoppm() -> str | None:
    return shutil.which("pdftoppm")


def render(deck: Path, out: Path) -> list[Path]:
    soffice = find_soffice()
    if soffice is None:
        raise ToolMissing("soffice not found; install LibreOffice with `brew install --cask libreoffice`")
    pdftoppm = find_pdftoppm()
    if pdftoppm is None:
        raise ToolMissing("pdftoppm not found; install poppler with `brew install poppler`")
    out.mkdir(parents=True, exist_ok=True)
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
    return sorted(pngs, key=lambda p: int(p.stem.split("-")[1]))


def _run(cmd: list[str]) -> None:
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        raise RenderError(f"{Path(cmd[0]).name} failed ({result.returncode}): {result.stderr.strip()}")
