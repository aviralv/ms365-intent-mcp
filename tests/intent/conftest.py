"""Shared pytest fixtures for intent tests.

``snapshot(name, schema)`` — golden JSON Schema snapshot fixture.

Write path (first run or SNAPSHOT_UPDATE=1):
    Writes ``tests/intent/snapshots/schemas/<name>.json`` with
    indent=2, sort_keys=True, and a trailing newline.

Assert path (subsequent runs):
    Reads the on-disk file and asserts byte-equality with the fresh schema.
    On mismatch the error message includes the exact update command.
"""

from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

import pytest

_SNAPSHOT_DIR = Path(__file__).parent / "snapshots" / "schemas"


def _serialise(schema: dict[str, Any]) -> bytes:
    """Stable serialisation — sorted keys, 2-space indent, trailing newline."""
    return (json.dumps(schema, indent=2, sort_keys=True) + "\n").encode()


@pytest.fixture
def snapshot():
    """Return a callable ``snapshot(name, schema)`` that writes or asserts."""

    def _snapshot(name: str, schema: dict[str, Any]) -> None:
        _SNAPSHOT_DIR.mkdir(parents=True, exist_ok=True)
        path = _SNAPSHOT_DIR / f"{name}.json"
        fresh = _serialise(schema)

        if not path.exists() or os.environ.get("SNAPSHOT_UPDATE") == "1":
            path.write_bytes(fresh)
            return

        on_disk = path.read_bytes()
        if on_disk != fresh:
            raise AssertionError(
                f"Schema snapshot mismatch for '{name}'.\n"
                f"Run the following command to update:\n\n"
                f"  SNAPSHOT_UPDATE=1 pytest tests/intent/test_schema_snapshots.py -k {name}\n\n"
                f"Diff (on-disk vs fresh):\n"
                f"  on-disk length: {len(on_disk)} bytes\n"
                f"  fresh length:   {len(fresh)} bytes"
            )

    return _snapshot
