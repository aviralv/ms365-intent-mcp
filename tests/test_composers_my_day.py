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


class TestMyDayIncludeBodies:
    def _client(self, body_html: str, join_url: str | None = None):
        from unittest.mock import AsyncMock

        captured = {}

        async def _get(endpoint, params=None, headers=None):
            if "calendarView" in endpoint:
                captured["select"] = (params or {}).get("$select", "")
                ev = {
                    "subject": "Sync",
                    "start": {"dateTime": "2026-08-05T09:00:00", "timeZone": "Europe/Berlin"},
                    "end": {"dateTime": "2026-08-05T09:30:00", "timeZone": "Europe/Berlin"},
                    "isOnlineMeeting": bool(join_url),
                    "onlineMeeting": {"joinUrl": join_url} if join_url else None,
                    "body": {"contentType": "html", "content": body_html},
                }
                return {"value": [ev]}
            return {"value": []}

        client = AsyncMock()
        client.get = AsyncMock(side_effect=_get)
        client.calendar_headers = lambda tz: {}
        return client, captured

    @pytest.mark.asyncio
    async def test_include_bodies_adds_select_and_populates(self, calendar_only_permissions):
        html = '<p>Agenda</p> <a href="https://wiki.example.com/x">Spec</a>'
        client, captured = self._client(html)
        data, _ = await compose_my_day(
            client, calendar_only_permissions, "2026-08-05", "Europe/Berlin",
            include_bodies=True,
        )
        assert "body" in captured["select"]
        ev = data["events"][0]
        assert "[Spec](https://wiki.example.com/x)" in ev["body"]
        assert ev["links"] == ["https://wiki.example.com/x"]

    @pytest.mark.asyncio
    async def test_default_omits_body(self, calendar_only_permissions):
        client, captured = self._client("<p>x</p>")
        data, _ = await compose_my_day(
            client, calendar_only_permissions, "2026-08-05", "Europe/Berlin",
        )
        assert "body" not in captured["select"]
        ev = data["events"][0]
        assert ev.get("body") is None
        assert ev.get("links", []) == []

    @pytest.mark.asyncio
    async def test_body_truncated_at_2000(self, calendar_only_permissions):
        html = "x" * 3000
        client, _ = self._client(html)
        data, _ = await compose_my_day(
            client, calendar_only_permissions, "2026-08-05", "Europe/Berlin",
            include_bodies=True,
        )
        assert len(data["events"][0]["body"]) == 2000

    @pytest.mark.asyncio
    async def test_join_url_excluded_link_before_truncation(self, calendar_only_permissions):
        join = "https://teams.microsoft.com/l/meetup-join/19:abc@thread.v2/0"
        html = ("z" * 2100) + '<a href="https://wiki.example.com/deep">Deep</a>' \
               + f'<a href="{join}">Join</a>'
        client, _ = self._client(html, join_url=join)
        data, _ = await compose_my_day(
            client, calendar_only_permissions, "2026-08-05", "Europe/Berlin",
            include_bodies=True,
        )
        ev = data["events"][0]
        assert ev["links"] == ["https://wiki.example.com/deep"]
        assert join not in ev["links"]

    @pytest.mark.asyncio
    async def test_boilerplate_only_body_collapses_to_none(self, calendar_only_permissions):
        """Issue #64: a no-agenda online meeting's body is pure Teams boilerplate
        and must collapse to None, not carry the join/dial-in noise."""
        sep = "_" * 80
        html = (
            f"<div>{sep}</div><div>Microsoft Teams meeting</div>"
            '<div>Join: <a href="https://teams.microsoft.com/meet/123">Click</a></div>'
            "<div>Meeting ID: 376 059 509 614 193</div>"
            "<div>Phone conference ID: 524 097 967#</div>"
            f"<div>{sep}</div>"
        )
        client, _ = self._client(html)
        data, _ = await compose_my_day(
            client, calendar_only_permissions, "2026-08-05", "Europe/Berlin",
            include_bodies=True,
        )
        assert data["events"][0]["body"] is None

    @pytest.mark.asyncio
    async def test_agenda_survives_boilerplate_strip(self, calendar_only_permissions):
        """Agenda text above the Teams separator is preserved; boilerplate below removed."""
        sep = "_" * 80
        html = (
            "<div>Agenda: review Q3 metrics</div>"
            f"<div>{sep}</div><div>Microsoft Teams meeting</div>"
            "<div>Meeting ID: 376 059 509 614 193</div>"
        )
        client, _ = self._client(html)
        data, _ = await compose_my_day(
            client, calendar_only_permissions, "2026-08-05", "Europe/Berlin",
            include_bodies=True,
        )
        body = data["events"][0]["body"]
        assert "Agenda: review Q3 metrics" in body
        assert "Microsoft Teams meeting" not in body
        assert "Meeting ID" not in body


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
