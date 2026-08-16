"""Unit tests for transcript composer helpers (issue #29 review hardening).

Covers the ``_dest_path`` path-traversal containment guard added after the
GPT-5 + Gemini review.
"""

from __future__ import annotations

import tempfile
from pathlib import Path
from unittest.mock import AsyncMock, patch

import pytest

from ms365_intent_mcp.composers.transcript import (
    RECAP_LINK_HINT,
    _dest_path,
    _unresolved_reason,
    compose_transcript,
)
from ms365_intent_mcp.permissions import PermissionRegistry
from ms365_intent_mcp.transcripts import Recording


@pytest.fixture
def full_permissions():
    return PermissionRegistry(["Sites.Read.All"])


def _rec(item_id: str, meeting_name: str, date_yyyymmdd: str) -> Recording:
    return Recording(
        name=f"{meeting_name}-{date_yyyymmdd}_120000-Meeting Recording.mp4",
        item_id=item_id,
        drive_id="drv",
        size=1,
        created=f"{date_yyyymmdd[:4]}-{date_yyyymmdd[4:6]}-{date_yyyymmdd[6:8]}T12:00:00Z",
        personal_site="https://sap-my.sharepoint.com/personal/u",
        web_url="https://sap-my.sharepoint.com/personal/u/x",
    )


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
    msg = _unresolved_reason(hint="Foo|2026-01-01", site_root="", drive_id="b!x", item_id="i")
    assert "site" in msg.lower() or "host" in msg.lower()


def test_unresolved_reason_falls_back_when_no_hint():
    msg = _unresolved_reason(hint="", site_root="", drive_id="", item_id="")
    assert msg  # non-empty; user-facing fallback


# ---------- list mode (ferret `list` parity) ----------


@pytest.mark.asyncio
async def test_list_mode_returns_date_ranked_recordings(full_permissions):
    """`list_recordings=True` runs discovery and returns recordings newest-first
    — the capability that surfaced the correct file in the #34 incident.

    Patches the three sub-discovery helpers (not the aggregate) so the real
    dedup + date-sort in ``_discover_all_recordings`` is exercised, and feeds
    them out-of-order to prove the ranking isn't just input order."""
    recs = [
        _rec("old", "Meeting with Bawa Kulkarni", "20260701"),
        _rec("new", "Call with Vaid, Aviral", "20260714"),
    ]
    vroom = AsyncMock()
    with (
        patch(
            "ms365_intent_mcp.composers.transcript._discover_own_drive",
            AsyncMock(return_value=recs),
        ),
        patch(
            "ms365_intent_mcp.composers.transcript._discover_search",
            AsyncMock(return_value=[]),
        ),
        patch(
            "ms365_intent_mcp.composers.transcript._discover_chats",
            AsyncMock(return_value=[]),
        ),
    ):
        data, markdown = await compose_transcript(
            AsyncMock(),
            vroom,
            full_permissions,
            url=None,
            name=None,
            item_id=None,
            drive_id=None,
            site_root=None,
            output_dir=None,
            list_recordings=True,
        )

    assert data["status"] == "ok"
    dates = [r["meeting_date"] for r in data["recordings"]]
    assert dates == ["2026-07-14", "2026-07-01"]  # newest first
    assert "Call with Vaid, Aviral" in markdown
    assert "2026-07-14" in markdown


@pytest.mark.asyncio
async def test_list_mode_empty_is_ok_not_error(full_permissions):
    with patch(
        "ms365_intent_mcp.composers.transcript._discover_all_recordings",
        AsyncMock(return_value=[]),
    ):
        data, _ = await compose_transcript(
            AsyncMock(),
            AsyncMock(),
            full_permissions,
            url=None,
            name=None,
            item_id=None,
            drive_id=None,
            site_root=None,
            output_dir=None,
            list_recordings=True,
        )
    assert data["status"] == "ok"
    assert data["recordings"] == []


# ---------- by-coords download path (#33) ----------


@pytest.mark.asyncio
async def test_by_coords_downloads_with_zero_discovery(full_permissions):
    """A known (site_root, drive_id, item_id) downloads directly — no discovery
    helpers are called (issue #33)."""
    vroom = AsyncMock()
    vroom.list_transcripts = AsyncMock(return_value=[{"id": "t1"}])
    vroom.download_transcript_to_file = AsyncMock(return_value=42)

    d = tempfile.mkdtemp()
    with patch(
        "ms365_intent_mcp.composers.transcript._resolve_from_name",
        AsyncMock(side_effect=AssertionError("discovery must not run")),
    ):
        data, _ = await compose_transcript(
            AsyncMock(),
            vroom,
            full_permissions,
            url=None,
            name=None,
            item_id="013KZ",
            drive_id="b!drv",
            site_root="https://sap-my.sharepoint.com/personal/u",
            output_dir=d,
            list_recordings=False,
        )

    assert data["status"] == "ok"
    assert data["byte_count"] == 42
    vroom.list_transcripts.assert_awaited_once_with(
        "https://sap-my.sharepoint.com/personal/u", "b!drv", "013KZ"
    )


# ---------- name multi-match surfaces alternatives (soft signal, #34) ----------


@pytest.mark.asyncio
async def test_name_multimatch_surfaces_alternatives_in_result(full_permissions):
    """When name matches several recordings, the freshest is downloaded but the
    response names the alternatives so a stale pick is never silent."""
    recs = [
        _rec("old", "Meeting with Bawa Kulkarni", "20260701"),
        _rec("new", "Meeting with Bawa Kulkarni", "20260714"),
    ]
    vroom = AsyncMock()
    vroom.list_transcripts = AsyncMock(return_value=[{"id": "t1"}])
    vroom.download_transcript_to_file = AsyncMock(return_value=10)

    d = tempfile.mkdtemp()
    with patch(
        "ms365_intent_mcp.composers.transcript._discover_all_recordings",
        AsyncMock(return_value=recs),
    ):
        data, markdown = await compose_transcript(
            AsyncMock(),
            vroom,
            full_permissions,
            url=None,
            name="Bawa",
            item_id=None,
            drive_id=None,
            site_root=None,
            output_dir=d,
            list_recordings=False,
        )

    assert data["status"] == "ok"
    assert data["meeting_date"] == "2026-07-14"  # freshest downloaded
    assert data["alternatives_count"] == 1
    assert "2026-07-01" in markdown  # the stale alternative is named


# ---------- colleague-hosted recording: recap-link hint (#43) ----------


@pytest.mark.asyncio
async def test_name_miss_surfaces_recap_link_hint(full_permissions):
    """A name that matches nothing points the caller at the url= recap-link
    workaround — the colleague-hosted-recording dead-end from #43."""
    with patch(
        "ms365_intent_mcp.composers.transcript._discover_all_recordings",
        AsyncMock(return_value=[]),
    ):
        data, markdown = await compose_transcript(
            AsyncMock(),
            AsyncMock(),
            full_permissions,
            url=None,
            name="DPDHL Regular Check-In",
            item_id=None,
            drive_id=None,
            site_root=None,
            output_dir=None,
            list_recordings=False,
        )

    assert data["status"] == "error"
    assert RECAP_LINK_HINT in data["message"]
    assert "url=" in markdown


@pytest.mark.asyncio
async def test_populated_list_footer_carries_recap_hint(full_permissions):
    """The #43 repro: list returns many recordings but not the target one (it's
    on a colleague's drive). The populated-list footer must still point at the
    recap-link workaround, not just the empty-list case."""
    recs = [_rec("a", "Some Other Meeting", "20260715")]
    with patch(
        "ms365_intent_mcp.composers.transcript._discover_all_recordings",
        AsyncMock(return_value=recs),
    ):
        data, markdown = await compose_transcript(
            AsyncMock(),
            AsyncMock(),
            full_permissions,
            url=None,
            name=None,
            item_id=None,
            drive_id=None,
            site_root=None,
            output_dir=None,
            list_recordings=True,
        )

    assert data["status"] == "ok"
    assert len(data["recordings"]) == 1
    assert RECAP_LINK_HINT in markdown


@pytest.mark.asyncio
async def test_empty_list_carries_recap_hint(full_permissions):
    """Empty discovery still surfaces the recap-link workaround."""
    with patch(
        "ms365_intent_mcp.composers.transcript._discover_all_recordings",
        AsyncMock(return_value=[]),
    ):
        _, markdown = await compose_transcript(
            AsyncMock(),
            AsyncMock(),
            full_permissions,
            url=None,
            name=None,
            item_id=None,
            drive_id=None,
            site_root=None,
            output_dir=None,
            list_recordings=True,
        )

    assert RECAP_LINK_HINT in markdown
