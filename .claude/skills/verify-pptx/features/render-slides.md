# Render slides

`deckcheck render` turns a deck into one PNG per slide. These images are the only proof of visual changes such as overflowing text, overlaps, or chart styling.

## Sub-features

- `render-png` writes `slide-1.png` to `slide-N.png` at 80 dpi, plus the intermediate PDF.
- `render-missing-tool` exits 3 with an install hint when soffice or pdftoppm is missing.

## How to get to it (user POV)

- Run `deckcheck render <deck.pptx> --out <dir>`.

## Driving it with deckcheck

Preconditions:

- Baseline preconditions from `README.md` hold.
- `doctor` prints a path for both `soffice` and `pdftoppm`. Install LibreOffice with `brew install --cask libreoffice` if soffice is missing.

- **Render.** Run `uv run --project deckcheck deckcheck render $RUN/decks/clean.pptx --out $RUN/render`. Exit 0. Stdout lists one PNG path per slide, and `$RUN/render` holds `clean.pdf` and `slide-1.png` to `slide-4.png`.
- **Inspect.** Open each changed slide's PNG and confirm that no text is clipped and no shapes overlap.
- **Missing tool.** On a machine without soffice, the same command exits 3 and stderr contains `brew install --cask libreoffice`. Report render as skipped.

## Gotchas

- LibreOffice is not PowerPoint. Font substitution and think-cell charts render differently. Treat a render as evidence of layout problems, not as pixel truth.
- `render` runs soffice with a private profile, so an open LibreOffice window does not block it.
