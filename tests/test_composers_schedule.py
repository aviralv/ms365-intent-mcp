"""Tests for schedule composer."""

from unittest.mock import AsyncMock

import pytest

from ms365_intent_mcp.composers.schedule import compose_schedule
from ms365_intent_mcp.graph import GraphAPIError
from ms365_intent_mcp.permissions import PermissionRegistry


@pytest.fixture
def full_permissions():
    return PermissionRegistry(["Calendars.ReadWrite"])


class TestScheduleBasic:
    @pytest.mark.asyncio
    async def test_returns_suggestions(self, full_permissions):
        client = AsyncMock()
        client.post = AsyncMock(return_value=_mock_find_meeting_times_response())

        result = await compose_schedule(
            client=client,
            permissions=full_permissions,
            attendees=[{"email": "bob@example.com", "name": "Bob"}],
            duration_minutes=30,
            constraints=None,
        )
        _, markdown = result
        assert "10:00" in markdown
        assert "100" in markdown

    @pytest.mark.asyncio
    async def test_no_suggestions_returns_helpful_message(self, full_permissions):
        client = AsyncMock()
        client.post = AsyncMock(
            return_value={
                "meetingTimeSuggestions": [],
                "emptySuggestionsReason": "AttendeesUnavailable",
            }
        )

        _, result = await compose_schedule(
            client=client,
            permissions=full_permissions,
            attendees=[{"email": "bob@example.com"}],
            duration_minutes=30,
            constraints=None,
        )
        assert "No available" in result or "unavailable" in result.lower()

    @pytest.mark.asyncio
    async def test_missing_calendar_scope_returns_message(self):
        client = AsyncMock()
        permissions = PermissionRegistry([])

        _, result = await compose_schedule(
            client=client,
            permissions=permissions,
            attendees=[{"email": "x@x.com"}],
            duration_minutes=30,
            constraints=None,
        )
        assert "Calendars.ReadWrite" in result

    @pytest.mark.asyncio
    async def test_graph_error_returns_error_section(self, full_permissions):
        client = AsyncMock()
        client.post = AsyncMock(side_effect=GraphAPIError(400, "InvalidRequest", "bad attendees"))

        _, result = await compose_schedule(
            client=client,
            permissions=full_permissions,
            attendees=[{"email": "bad"}],
            duration_minutes=30,
            constraints=None,
        )
        assert "⚠️" in result or "bad attendees" in result


def _mock_find_meeting_times_response():
    return {
        "meetingTimeSuggestions": [
            {
                "meetingTimeSlot": {
                    "start": {"dateTime": "2026-05-20T10:00:00", "timeZone": "UTC"},
                    "end": {"dateTime": "2026-05-20T10:30:00", "timeZone": "UTC"},
                },
                "confidence": 100.0,
                "attendeeAvailability": [],
            }
        ]
    }


class TestScheduleTimezoneAware:
    @pytest.mark.asyncio
    async def test_schedule_slot_is_offset_aware_with_tz(self):
        """Slots must carry offset-aware ISO strings and sibling tz keys."""
        client = AsyncMock()
        client.post = AsyncMock(
            return_value={
                "meetingTimeSuggestions": [
                    {
                        "meetingTimeSlot": {
                            "start": {
                                "dateTime": "2026-07-29T14:00:00.0000000",
                                "timeZone": "UTC",
                            },
                            "end": {
                                "dateTime": "2026-07-29T14:30:00.0000000",
                                "timeZone": "UTC",
                            },
                        },
                        "confidence": 90.0,
                        "attendeeAvailability": [],
                    }
                ]
            }
        )
        permissions = PermissionRegistry(["Calendars.ReadWrite"])

        data, _markdown = await compose_schedule(
            client=client,
            permissions=permissions,
            attendees=[{"email": "bob@example.com", "name": "Bob"}],
            duration_minutes=30,
            constraints=None,
        )

        slot = data["suggestions"][0]
        assert slot["start"] == "2026-07-29T14:00:00+00:00"
        assert slot["end"] == "2026-07-29T14:30:00+00:00"
        assert slot["start_timezone"] == "UTC"
        assert slot["end_timezone"] == "UTC"
        assert slot["confidence"] == 0.9
