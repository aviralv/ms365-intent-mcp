"""Unit tests for _schedule_impl — mocked context, no FastMCP."""

from unittest.mock import AsyncMock, MagicMock

import pytest

from ms365_intent_mcp.graph import GraphAPIError
from ms365_intent_mcp.intent._shared import ErrorResponse
from ms365_intent_mcp.intent.schedule.impl import _schedule_impl
from ms365_intent_mcp.intent.schedule.schemas import SchedulePayload, ScheduleSuggestions


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


class TestScheduleV1HappyPath:
    @pytest.mark.asyncio
    async def test_single_attendee_returns_typed_response(self, monkeypatch):
        ctx, _, _ = _mock_ctx()

        async def _fake(client, permissions, attendees, duration_minutes, constraints):
            return {"suggestions": []}, "## Meeting Suggestions\n- Monday 10am"

        monkeypatch.setattr(
            "ms365_intent_mcp.intent.schedule.impl.compose_schedule",
            _fake,
        )

        payload = SchedulePayload.model_validate(
            {
                "attendees": [{"email": "alice@example.com"}],
                "duration_minutes": 30,
            }
        )
        response = await _schedule_impl(ctx, payload)

        assert isinstance(response, ScheduleSuggestions)
        assert response.type == "schedule_suggestions"
        assert "Meeting Suggestions" in response.rendered_markdown
        assert response.suggestions == []

    @pytest.mark.asyncio
    async def test_attendee_name_falls_back_to_email(self, monkeypatch):
        """Attendee without name should use email as the name in the flat dict."""
        ctx, _, _ = _mock_ctx()
        captured_attendees: list = []

        async def _fake(client, permissions, attendees, duration_minutes, constraints):
            captured_attendees.extend(attendees)
            return {"suggestions": []}, "suggestions"

        monkeypatch.setattr(
            "ms365_intent_mcp.intent.schedule.impl.compose_schedule",
            _fake,
        )

        payload = SchedulePayload.model_validate(
            {
                "attendees": [{"email": "bob@example.com"}],
            }
        )
        await _schedule_impl(ctx, payload)

        assert len(captured_attendees) == 1
        assert captured_attendees[0]["email"] == "bob@example.com"
        assert captured_attendees[0]["name"] == "bob@example.com"


class TestScheduleV1AttendeeConversion:
    @pytest.mark.asyncio
    async def test_typed_attendee_converted_to_flat_dict(self, monkeypatch):
        """list[Attendee] must be converted to flat dict[email, name] for composer."""
        ctx, _, _ = _mock_ctx()
        captured_attendees: list = []

        async def _fake(client, permissions, attendees, duration_minutes, constraints):
            captured_attendees.extend(attendees)
            return {"suggestions": []}, "suggestions"

        monkeypatch.setattr(
            "ms365_intent_mcp.intent.schedule.impl.compose_schedule",
            _fake,
        )

        payload = SchedulePayload.model_validate(
            {
                "attendees": [
                    {"email": "alice@example.com", "name": "Alice"},
                    {"email": "bob@example.com"},
                ],
                "duration_minutes": 45,
            }
        )
        await _schedule_impl(ctx, payload)

        assert len(captured_attendees) == 2
        assert captured_attendees[0] == {"email": "alice@example.com", "name": "Alice"}
        assert captured_attendees[1] == {"email": "bob@example.com", "name": "bob@example.com"}


class TestScheduleV1Validation:
    def test_empty_attendees_rejected(self):
        """min_length=1 must reject empty attendee list."""
        import pydantic

        with pytest.raises(pydantic.ValidationError) as exc_info:
            SchedulePayload.model_validate(
                {
                    "attendees": [],
                    "duration_minutes": 30,
                }
            )
        errors = exc_info.value.errors()
        assert any("attendees" in str(e) for e in errors)

    def test_duration_below_minimum_rejected(self):
        import pydantic

        with pytest.raises(pydantic.ValidationError):
            SchedulePayload.model_validate(
                {
                    "attendees": [{"email": "alice@example.com"}],
                    "duration_minutes": 4,
                }
            )

    def test_duration_above_maximum_rejected(self):
        import pydantic

        with pytest.raises(pydantic.ValidationError):
            SchedulePayload.model_validate(
                {
                    "attendees": [{"email": "alice@example.com"}],
                    "duration_minutes": 481,
                }
            )


class TestScheduleV1ErrorHandling:
    @pytest.mark.asyncio
    async def test_graph_api_error_returns_error_response(self, monkeypatch):
        ctx, _, _ = _mock_ctx()

        async def _fail(client, permissions, attendees, duration_minutes, constraints):
            raise GraphAPIError(
                status_code=429, error_code="TooManyRequests", message="rate limited"
            )

        monkeypatch.setattr(
            "ms365_intent_mcp.intent.schedule.impl.compose_schedule",
            _fail,
        )

        payload = SchedulePayload.model_validate(
            {
                "attendees": [{"email": "alice@example.com"}],
            }
        )
        response = await _schedule_impl(ctx, payload)

        assert isinstance(response, ErrorResponse)
        assert response.type == "error"
        assert response.code == "rate_limited"
        assert response.retryable is True
