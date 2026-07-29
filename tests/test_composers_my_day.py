"""Tests for my_day composer — partial failure and formatting."""

import asyncio
from unittest.mock import AsyncMock

import pytest

from ms365_intent_mcp.composers.my_day import compose_my_day
from ms365_intent_mcp.graph import GraphAPIError
from ms365_intent_mcp.permissions import PermissionRegistry


@pytest.fixture
def full_permissions():
    return PermissionRegistry([
        "Calendars.ReadWrite", "Mail.Read", "Chat.ReadWrite",
    ])


@pytest.fixture
def calendar_only_permissions():
    return PermissionRegistry(["Calendars.ReadWrite"])


class TestMyDayAllSucceed:
    @pytest.mark.asyncio
    async def test_returns_all_sections(self, full_permissions):
        client = AsyncMock()
        client.get = AsyncMock(side_effect=_mock_graph_get)
        client.calendar_headers = lambda tz: {"Prefer": f'outlook.timezone="{tz}"'}

        result = await compose_my_day(
            client=client,
            permissions=full_permissions,
            date="2026-05-15",
            timezone="Europe/Berlin",
        )
        _, markdown = result
        assert "Calendar" in markdown or "Standup" in markdown
        assert "Mail" in markdown
        assert "Teams" in markdown


class TestMyDayPartialFailure:
    @pytest.mark.asyncio
    async def test_mail_failure_still_returns_calendar(self, full_permissions):
        client = AsyncMock()

        async def _get(endpoint, params=None, headers=None):
            if "mailFolders" in endpoint or "messages" in endpoint:
                raise GraphAPIError(500, "InternalError", "mail is down")
            if "calendarView" in endpoint:
                return _mock_calendar_response()
            if "chats" in endpoint:
                return {"value": []}
            return {"value": []}

        client.get = AsyncMock(side_effect=_get)
        client.calendar_headers = lambda tz: {"Prefer": f'outlook.timezone="{tz}"'}

        _, result = await compose_my_day(
            client=client,
            permissions=full_permissions,
            date="2026-05-15",
            timezone="Europe/Berlin",
        )
        assert "⚠️" in result
        assert "Standup" in result


class TestMyDayMissingPermissions:
    @pytest.mark.asyncio
    async def test_no_chat_scope_skips_teams(self, calendar_only_permissions):
        client = AsyncMock()

        async def _get(endpoint, params=None, headers=None):
            if "calendarView" in endpoint:
                return _mock_calendar_response()
            if "mailFolders" in endpoint:
                return {"unreadItemCount": 0}
            if "messages" in endpoint:
                return {"value": []}
            return {"value": []}

        client.get = AsyncMock(side_effect=_get)
        client.calendar_headers = lambda tz: {"Prefer": f'outlook.timezone="{tz}"'}

        _, result = await compose_my_day(
            client=client,
            permissions=calendar_only_permissions,
            date="2026-05-15",
            timezone="Europe/Berlin",
        )
        assert "ℹ️" in result
        assert "Chat.ReadWrite" in result


def _mock_calendar_response():
    return {"value": [{
        "subject": "Standup",
        "start": {"dateTime": "2026-05-15T09:00:00"},
        "end": {"dateTime": "2026-05-15T09:30:00"},
        "location": {"displayName": ""},
        "isOnlineMeeting": True,
        "attendees": [],
        "organizer": {"emailAddress": {"name": "Alice"}},
    }]}


async def _mock_graph_get(endpoint, params=None, headers=None):
    if "calendarView" in endpoint:
        return _mock_calendar_response()
    if "mailFolders/inbox" in endpoint and "messages" not in endpoint:
        return {"unreadItemCount": 3}
    if "messages" in endpoint:
        return {"value": []}
    if "chats" in endpoint:
        return {"value": []}
    return {"value": []}


class TestMyDayEventTimezones:
    @pytest.mark.asyncio
    async def test_event_start_is_offset_aware_with_tz_sibling(self, full_permissions):
        client = AsyncMock()

        async def fake_get(endpoint, params=None, headers=None):
            if "calendarView" in endpoint:
                return {"value": [
                    {
                        "subject": "Standup",
                        "start": {"dateTime": "2026-07-29T14:00:00.0000000", "timeZone": "UTC"},
                        "end": {"dateTime": "2026-07-29T14:30:00.0000000", "timeZone": "UTC"},
                        "location": {"displayName": ""},
                        "isOnlineMeeting": True,
                    },
                    {
                        "subject": "Offsite",
                        "isAllDay": True,
                        "start": {"dateTime": "2026-07-30", "timeZone": "UTC"},
                        "end": {"dateTime": "2026-07-31", "timeZone": "UTC"},
                        "location": {"displayName": ""},
                        "isOnlineMeeting": False,
                    },
                ]}
            if "mailFolders/inbox" in endpoint and "messages" not in endpoint:
                return {"unreadItemCount": 0}
            return {"value": []}

        client.get = AsyncMock(side_effect=fake_get)
        client.calendar_headers = lambda tz: {"Prefer": f'outlook.timezone="{tz}"'}

        data, _ = await compose_my_day(
            client=client,
            permissions=full_permissions,
            date="2026-07-29",
            timezone="UTC",
        )

        ev = data["events"][0]
        assert ev["start"] == "2026-07-29T14:00:00+00:00"
        assert ev["start_timezone"] == "UTC"
        assert ev["end_timezone"] == "UTC"

        allday = data["events"][1]
        assert allday["start"] == "2026-07-30"
        assert allday["start_timezone"] == "UTC"
