"""Unit tests for transcript composer helpers (issue #29 review hardening).

Covers the ``_dest_path`` path-traversal containment guard added after the
GPT-5 + Gemini review.
"""

from __future__ import annotations

import tempfile
from pathlib import Path

import pytest

from ms365_intent_mcp.composers.transcript import _dest_path, _unresolved_reason


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


# ---------- _unresolved_reason: opaque-error fix (issue #31 ask #1) ----------


def test_unresolved_reason_names_missing_drive_id():
    """When the recording resolved a name+item but no drive id, the error must
    say so — not echo the meeting name (the old opaque `hint`)."""
    msg = _unresolved_reason(
        hint="Call with Georgi, Dimitri|2026-07-13",
        site_root="https://sap-my.sharepoint.com/personal/u",
        drive_id="",
        item_id="013KZ...",
    )
    assert "drive" in msg.lower()
    assert "Call with Georgi, Dimitri" in msg  # still names the recording


def test_unresolved_reason_names_missing_item_id():
    msg = _unresolved_reason(
        hint="Foo|2026-01-01",
        site_root="https://sap-my.sharepoint.com/personal/u",
        drive_id="b!x",
        item_id="",
    )
    assert "item" in msg.lower()


def test_unresolved_reason_names_missing_site_root():
    msg = _unresolved_reason(
        hint="Foo|2026-01-01", site_root="", drive_id="b!x", item_id="i"
    )
    assert "site" in msg.lower() or "host" in msg.lower()


def test_unresolved_reason_falls_back_when_no_hint():
    msg = _unresolved_reason(hint="", site_root="", drive_id="", item_id="")
    assert msg  # non-empty; user-facing fallback
