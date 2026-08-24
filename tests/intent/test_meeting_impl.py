"""Unit tests for _meeting_impl — mocked context, no FastMCP."""

from unittest.mock import AsyncMock, MagicMock

import pytest

from ms365_intent_mcp.graph import GraphAPIError
from ms365_intent_mcp.intent._shared import ErrorResponse
from ms365_intent_mcp.intent.meeting.impl import _meeting_impl
from ms365_intent_mcp.intent.meeting.schemas import (
    MeetingDetail,
    MeetingPayload,
)


def _mock_ctx():
    """Build a mocked FastMCP Context with the three deps the impls need."""
    ctx = MagicMock()
    config = MagicMock(default_timezone="Europe/Berlin")
    client = AsyncMock()
    permissions = MagicMock()
    permissions.check = MagicMock(return_value=None)
    ctx.request_context.lifespan_context = {
        "config": config,
        "client": client,
        "permissions": permissions,
    }
    return ctx, client, permissions


class TestMeetingV1HappyPath:
    @pytest.mark.asyncio
    async def test_returns_typed_meeting_detail(self, monkeypatch):
        ctx, _, _ = _mock_ctx()

        async def _fake_compose(client, permissions, identifier, timezone):
            return {
                "id": "event-123",
                "subject": "Weekly Sync",
                "start": "2026-07-08T10:00:00",
                "end": "2026-07-08T11:00:00",
                "organizer": {"name": "Alice", "email": "alice@example.com"},
                "attendees": [],
                "location": None,
                "online_meeting": None,
                "recording": None,
            }, "## Weekly Sync\n\nOrganizer: Alice"

        monkeypatch.setattr(
            "ms365_intent_mcp.intent.meeting.impl.compose_meeting",
            _fake_compose,
        )

        payload = MeetingPayload(identifier="Weekly Sync")
        response = await _meeting_impl(ctx, payload)

        assert isinstance(response, MeetingDetail)
        assert response.type == "meeting_detail"
        assert "Weekly Sync" in response.rendered_markdown

    @pytest.mark.asyncio
    async def test_identifier_next_passed_through(self, monkeypatch):
        ctx, _, _ = _mock_ctx()
        captured = []

        async def _fake_compose(client, permissions, identifier, timezone):
            captured.append(identifier)
            return {
                "id": "evt-next",
                "subject": "Next Meeting",
                "start": "2026-07-08T14:00:00",
                "end": "2026-07-08T14:30:00",
                "organizer": {"name": "Bob"},
                "attendees": [],
                "location": None,
                "online_meeting": None,
                "recording": None,
            }, "## Next Meeting"

        monkeypatch.setattr(
            "ms365_intent_mcp.intent.meeting.impl.compose_meeting",
            _fake_compose,
        )

        payload = MeetingPayload(identifier="next")
        await _meeting_impl(ctx, payload)

        assert captured == ["next"], "identifier 'next' should be passed unchanged"

    @pytest.mark.asyncio
    async def test_structured_fields_populated(self, monkeypatch):
        """Structured fields are populated from the composer's dict."""
        ctx, _, _ = _mock_ctx()

        async def _fake_compose(client, permissions, identifier, timezone):
            return {
                "id": "evt-standup",
                "subject": "Standup",
                "start": "2026-07-08T09:00:00",
                "end": "2026-07-08T09:15:00",
                "organizer": {"name": "Carol", "email": None},
                "attendees": [],
                "location": None,
                "online_meeting": None,
                "recording": None,
            }, "## Standup\n"

        monkeypatch.setattr(
            "ms365_intent_mcp.intent.meeting.impl.compose_meeting",
            _fake_compose,
        )

        payload = MeetingPayload(identifier="Standup")
        response = await _meeting_impl(ctx, payload)

        assert response.id == "evt-standup"
        assert response.organizer.name == "Carol"
        assert response.attendees == []

    @pytest.mark.asyncio
    async def test_recording_with_occurrence_date_fields_survives_validation(self, monkeypatch):
        """Regression: the recording dict now carries recording_date /
        occurrence_date / date_matches_occurrence. These must be in the
        RecordingMetadata schema (extra='forbid'), else model_validate raises
        and the recording is silently swallowed to None (issue #79/#51)."""
        ctx, _, _ = _mock_ctx()

        async def _fake_compose(client, permissions, identifier, timezone):
            return {
                "id": "evt-recurring",
                "subject": "Catch-up",
                "start": "2026-08-24T09:00:00",
                "end": "2026-08-24T09:30:00",
                "organizer": {"name": "Avi", "email": "avi@example.com"},
                "attendees": [],
                "location": None,
                "online_meeting": None,
                "recording": {
                    "recording_url": "https://sap-my.sharepoint.com/:v:/p/avi/IQ",
                    "display_name": "Catch-up-20260518_110724-Meeting Recording.mp4",
                    "transcript_ready": True,
                    "drive_id": None,
                    "drive_item_id": None,
                    "owner_upn": None,
                    "vroom_url": None,
                    "recording_date": "2026-05-18",
                    "occurrence_date": "2026-08-24",
                    "date_matches_occurrence": False,
                },
            }, "## Catch-up\n"

        monkeypatch.setattr(
            "ms365_intent_mcp.intent.meeting.impl.compose_meeting",
            _fake_compose,
        )

        payload = MeetingPayload(identifier="Catch-up")
        response = await _meeting_impl(ctx, payload)

        assert response.recording is not None, "recording must not be swallowed to None"
        assert response.recording.recording_date == "2026-05-18"
        assert response.recording.occurrence_date == "2026-08-24"
        assert response.recording.date_matches_occurrence is False


class TestMeetingV1PayloadValidation:
    def test_empty_identifier_rejected(self):
        """MeetingPayload enforces min_length=1 on identifier."""
        import pytest as pt

        with pt.raises(Exception):
            MeetingPayload(identifier="")

    def test_valid_identifier_accepted(self):
        payload = MeetingPayload(identifier="any-event-id-or-subject")
        assert payload.identifier == "any-event-id-or-subject"


class TestMeetingV1ErrorHandling:
    @pytest.mark.asyncio
    async def test_graph_api_error_returns_error_response(self, monkeypatch):
        ctx, _, _ = _mock_ctx()

        async def _failing_compose(client, permissions, identifier, timezone):
            raise GraphAPIError(status_code=500, error_code="InternalServerError", message="boom")

        monkeypatch.setattr(
            "ms365_intent_mcp.intent.meeting.impl.compose_meeting",
            _failing_compose,
        )

        payload = MeetingPayload(identifier="some-meeting")
        response = await _meeting_impl(ctx, payload)

        assert isinstance(response, ErrorResponse)
        assert response.type == "error"
        assert response.code == "graph_api_error"
        assert "boom" in response.message
