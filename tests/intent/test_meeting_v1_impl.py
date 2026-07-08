"""Unit tests for _meeting_v1_impl — mocked context, no FastMCP."""

from unittest.mock import AsyncMock, MagicMock

import pytest

from ms365_intent_mcp.intent.meeting.impl import _meeting_v1_impl
from ms365_intent_mcp.intent.meeting.schemas import (
    MeetingDetail,
    MeetingPayload,
)
from ms365_intent_mcp.intent._shared import ErrorResponse
from ms365_intent_mcp.graph import GraphAPIError


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
            return "## Weekly Sync\n\nOrganizer: Alice"

        monkeypatch.setattr(
            "ms365_intent_mcp.intent.meeting.impl.compose_meeting",
            _fake_compose,
        )

        payload = MeetingPayload(identifier="Weekly Sync")
        response = await _meeting_v1_impl(ctx, payload)

        assert isinstance(response, MeetingDetail)
        assert response.type == "meeting_detail"
        assert "Weekly Sync" in response.rendered_markdown

    @pytest.mark.asyncio
    async def test_identifier_next_passed_through(self, monkeypatch):
        ctx, _, _ = _mock_ctx()
        captured = []

        async def _fake_compose(client, permissions, identifier, timezone):
            captured.append(identifier)
            return "## Next Meeting"

        monkeypatch.setattr(
            "ms365_intent_mcp.intent.meeting.impl.compose_meeting",
            _fake_compose,
        )

        payload = MeetingPayload(identifier="next")
        await _meeting_v1_impl(ctx, payload)

        assert captured == ["next"], "identifier 'next' should be passed unchanged"

    @pytest.mark.asyncio
    async def test_placeholder_fields_present(self, monkeypatch):
        """Structured fields are stubs — verify they're present with placeholder values."""
        ctx, _, _ = _mock_ctx()

        async def _fake_compose(client, permissions, identifier, timezone):
            return "## Standup\n"

        monkeypatch.setattr(
            "ms365_intent_mcp.intent.meeting.impl.compose_meeting",
            _fake_compose,
        )

        payload = MeetingPayload(identifier="Standup")
        response = await _meeting_v1_impl(ctx, payload)

        assert response.id == "pending-composer-dict-refactor"
        assert response.organizer.name == "unknown"
        assert response.attendees == []


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
        response = await _meeting_v1_impl(ctx, payload)

        assert isinstance(response, ErrorResponse)
        assert response.type == "error"
        assert response.code == "graph_api_error"
        assert "boom" in response.message
