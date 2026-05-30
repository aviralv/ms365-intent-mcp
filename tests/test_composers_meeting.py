"""Tests for meeting composer."""

from unittest.mock import AsyncMock

import pytest

from ms365_intent_mcp.composers.meeting import compose_meeting
from ms365_intent_mcp.graph import GraphAPIError
from ms365_intent_mcp.permissions import PermissionRegistry


@pytest.fixture
def full_permissions():
    return PermissionRegistry(["Calendars.ReadWrite", "Chat.ReadWrite"])


class TestMeetingById:
    @pytest.mark.asyncio
    async def test_returns_event_details(self, full_permissions):
        client = AsyncMock()
        client.get = AsyncMock(return_value=_full_event())

        result = await compose_meeting(
            client=client,
            permissions=full_permissions,
            identifier="event-id-123",
            timezone="Europe/Berlin",
        )
        assert "Team Sync" in result
        assert "Alice" in result


class TestMeetingBySubject:
    @pytest.mark.asyncio
    async def test_searches_by_subject(self, full_permissions):
        client = AsyncMock()

        async def _get(endpoint, params=None, headers=None):
            if "calendarView" in endpoint:
                return {"value": [_full_event()]}
            return _full_event()

        client.get = AsyncMock(side_effect=_get)
        result = await compose_meeting(
            client=client,
            permissions=full_permissions,
            identifier="Team Sync",
            timezone="Europe/Berlin",
        )
        assert "Team Sync" in result


class TestMeetingNext:
    @pytest.mark.asyncio
    async def test_next_finds_upcoming(self, full_permissions):
        client = AsyncMock()
        client.get = AsyncMock(return_value={"value": [_full_event()]})

        result = await compose_meeting(
            client=client,
            permissions=full_permissions,
            identifier="next",
            timezone="Europe/Berlin",
        )
        assert "Team Sync" in result


def _full_event():
    return {
        "id": "event-id-123",
        "subject": "Team Sync",
        "start": {"dateTime": "2026-05-15T14:00:00"},
        "end": {"dateTime": "2026-05-15T14:30:00"},
        "location": {"displayName": "Room B"},
        "isOnlineMeeting": True,
        "onlineMeeting": {"joinUrl": "https://teams.microsoft.com/l/meetup-join/123"},
        "organizer": {"emailAddress": {"name": "Alice", "address": "alice@example.com"}},
        "attendees": [
            {"emailAddress": {"name": "Bob", "address": "bob@example.com"}, "status": {"response": "accepted"}},
        ],
        "body": {"content": "<p>Agenda: review Q2 progress</p>", "contentType": "html"},
    }
