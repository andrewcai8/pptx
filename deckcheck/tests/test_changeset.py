from __future__ import annotations

from deckcheck.changeset.model import SCHEMA_DIR, schemas


def test_committed_schemas_match_the_models() -> None:
    for name, text in schemas().items():
        assert (SCHEMA_DIR / name).read_text() == text, (
            f"{name} is stale; run uv run --project deckcheck python -m deckcheck.changeset.model"
        )
