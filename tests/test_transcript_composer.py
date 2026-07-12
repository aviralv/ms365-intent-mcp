"""Unit tests for transcript composer helpers (issue #29 review hardening).

Covers the ``_dest_path`` path-traversal containment guard added after the
GPT-5 + Gemini review.
"""

from __future__ import annotations

import tempfile
from pathlib import Path

import pytest

from ms365_intent_mcp.composers.transcript import _dest_path


def test_dest_path_writes_under_base():
    d = tempfile.mkdtemp()
    p = _dest_path(d, "Sprint Review|2026-07-09", "abc123def456")
    assert p.parent == Path(d).resolve()
    assert p.name == "Sprint_Review-2026-07-09-abc123de.vtt"


def test_dest_path_neutralizes_traversal_in_name():
    """A meeting name / id carrying `../` must not escape the output dir. The
    sanitizer strips path separators, so even with `..` fragments the file
    resolves to a plain name directly under base (no traversal)."""
    d = tempfile.mkdtemp()
    p = _dest_path(d, "../../etc/passwd|", "x/../y")
    assert p.parent == Path(d).resolve()  # the real guarantee: stays under base
    assert "/" not in p.name  # no separators survived sanitization


def test_dest_path_default_dir_when_none():
    p = _dest_path(None, "Foo|2026-01-01", "deadbeef")
    assert p.parent == (Path.home() / ".cache" / "ms365-intent-mcp" / "transcripts").resolve()


def test_dest_path_empty_name_falls_back_to_transcript():
    d = tempfile.mkdtemp()
    p = _dest_path(d, "", "abcd1234")
    assert p.name.startswith("transcript-")
