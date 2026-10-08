---
name: verify-pptx
description: Prove a generated or edited PowerPoint deck is ready to ship by running deckcheck (house-style check, slide diff against the source deck, slide render) and capturing evidence. Use after producing or changing any .pptx in this repo, before claiming a deck is one-shot, or when deckcheck or standards/house-style.yaml changes.
---

# Verify a deck

The surface is the `deckcheck` CLI in `deckcheck/`. It reads `standards/house-style.yaml`, the company slide standard. There is no server and no web UI yet. When the review site or the OneDrive publish step lands, add them to `features/` with `/maintain-verification-skill`.

Read `features/README.md` before driving, then follow the feature file that matches the claim you are proving.

## Launch

Run from the repo root. Install once per checkout.

```bash
uv sync --project deckcheck
```

Each command is a short-lived process, so there is nothing to keep alive. Give every run its own evidence directory, and parallel runs never collide:

```bash
RUN=artifacts/verify-pptx/$(date +%Y%m%d-%H%M%S)-$$
```

## Doctor

```bash
uv run --project deckcheck deckcheck doctor
```

Require exit 0, `rules: <repo>/standards/house-style.yaml`, and a `rule ids:` line. `soffice: missing` means `render` is unavailable. Report a render step as skipped in that case. Do not count it as passed.

Then run the self-test. It proves deckcheck still catches every rule on known decks before you trust it on a real one:

```bash
.claude/skills/verify-pptx/scripts/selftest.sh
```

It prints `SELFTEST PASS` and the evidence path, or `SELFTEST FAIL: <reason>` and exits 1. Set `RUN_ID=<name>` to choose the directory name.

A change to `standards/house-style.yaml` or `deckcheck/src/deckcheck/rules.py` also needs the known-good corpus. It proves real BCG decks still pass:

```bash
uv run --project deckcheck python .claude/skills/verify-pptx/scripts/corpus.py
```

Require `SELFTEST PASS` and `CORPUS PASS` both. `CORPUS INCOMPLETE` means a deck was unreachable. Report the corpus as not run, never as passed.

## Drive

To prove a deck that the pipeline produced from a source deck:

1. Record the source hash before the pipeline runs. Run `shasum -a 256 <source.pptx> > $RUN/source.sha256`.
2. Run `uv run --project deckcheck deckcheck check <new.pptx> --out $RUN/check`. Exit 0 prints `PASS`. Exit 1 prints `FAIL` and one `slide N [rule-id] shape: message | evidence` line per violation.
3. Run `uv run --project deckcheck deckcheck diff <source.pptx> <new.pptx> --out $RUN/diff`. Every slide listed as `changed`, `added`, or `removed` must be one that the request asked for.
4. Run `shasum -a 256 -c $RUN/source.sha256`. It must print `OK`. The pipeline writes new files and never edits the source.
5. If `doctor` found soffice, run `uv run --project deckcheck deckcheck render <new.pptx> --out $RUN/render`, then open the changed slides' PNGs and look at them.

Exit codes for every subcommand are `0` ok, `1` violations, `2` bad input or config, and `3` missing external tool.

Never point deckcheck at a file inside a synced OneDrive folder while proving a change. Copy the file into `$RUN/` first.

## Evidence

Everything goes under `artifacts/verify-pptx/<run>/`. Git ignores that directory, and cleanup never touches it.

- `check/report.json` holds the deck sha256, the rule ids applied, `passed`, and every violation.
- `check/outline.md` lists each slide's title, its text, and the resolved fonts per paragraph. Read it to confirm the content, not just the style.
- `diff/diff.md` and `diff/diff.json` hold both hashes and the per-slide status with a text diff.
- `render/slide-N.png` holds the slide images when soffice exists.

Proof standards:

- Check the real output file the pipeline wrote. A hand-built copy or a deck rebuilt in the test does not count.
- A deck is one-shot only when `check` passes, `diff` is scoped to the requested slides, and the source hash is unchanged. All three, every time.
- `diff` compares text only. A font, position, or color change shows as `unchanged`. Prove visual changes with `render`, or say that the visual change is unverified.
- A rule you disabled or loosened in `standards/house-style.yaml` to get a pass is a failed proof. Report the violation instead.

## Cleanup

deckcheck starts no processes. `selftest.sh` deletes its own scratch directory on exit. Remove only scratch copies you made outside `$RUN/`. Never delete `artifacts/verify-pptx/`.

## Helpers

- `scripts/selftest.sh` builds sample decks with `deckcheck/scripts/make_sample_decks.py`. It asserts that `clean.pptx` passes, that `dirty.pptx` fails with exactly the expected `(slide, rule)` pairs, and that the diff touches only slides 2 and 5.
- `uv run --project deckcheck python deckcheck/scripts/make_sample_decks.py <dir>` writes `clean.pptx`, `dirty.pptx`, and `clean-v2.pptx`, and prints the expected violations for `dirty.pptx`.
- `cd deckcheck && uv run pytest -q` runs the unit tests.
- `scripts/corpus.py [RUN_DIR]` fetches the decks in `corpus/known-good.yaml` into `artifacts/verify-pptx/corpus-cache/`, checks each one, and fails on any violation without a waiver or any waiver that no longer fires. See `features/known-good-corpus.md`.
