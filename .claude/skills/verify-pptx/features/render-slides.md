# Render slides

`deckcheck render` turns a deck into one PNG per slide. These images are the only proof of visual changes such as overflowing text, overlaps, or chart styling.

## Sub-features

- `render-png` writes `slide-1.png` to `slide-N.png` at 80 dpi, plus the intermediate PDF.
- `render-fonts` writes `fonts.json`, which maps each font the deck uses to the family `fc-match` found, with `substituted` true or false. It prints one `substituted: <font> -> <family>` line per substituted font after the PNG paths. A metric-compatible clone, such as Carlito for Calibri, is not a substitution.
- `render-missing-tool` exits 3 with an install hint when fc-match, soffice, or pdftoppm is missing. It checks fc-match first and writes nothing.

## How to get to it (user POV)

- Run `deckcheck render <deck.pptx> --out <dir>`.

## Driving it with deckcheck

Preconditions:

- Baseline preconditions from `README.md` hold.
- `doctor` prints a path for both `soffice` and `pdftoppm`. Install LibreOffice with `brew install --cask libreoffice` if soffice is missing.
- `fc-match` exists. Install fontconfig with `brew install fontconfig` if it is missing.

- **Render.** Run `uv run --project deckcheck deckcheck render $RUN/decks/clean.pptx --out $RUN/render`. Exit 0. Stdout lists one PNG path per slide, and `$RUN/render` holds `clean.pdf`, `fonts.json`, and `slide-1.png` to `slide-4.png`.
- **Inspect.** Open each changed slide's PNG and confirm that no text is clipped and no shapes overlap.
- **Missing tool.** On a machine without soffice, the same command exits 3 and stderr contains `brew install --cask libreoffice`. Report render as skipped.

## Gotchas

- LibreOffice is not PowerPoint. Substituted fonts and think-cell charts render differently. Read `fonts.json` before you trust a wrap or an edge. Treat a render as evidence of layout problems, not as pixel truth.
- `render` runs soffice with a private profile, so an open LibreOffice window does not block it.
