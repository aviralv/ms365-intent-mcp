"""Tests for whats_new composer."""

from unittest.mock import AsyncMock

import pytest

from ms365_intent_mcp.composers.whats_new import compose_whats_new
from ms365_intent_mcp.graph import GraphAPIError
from ms365_intent_mcp.permissions import PermissionRegistry


@pytest.fixture
def full_permissions():
    return PermissionRegistry(["Calendars.ReadWrite", "Mail.Read", "Chat.ReadWrite"])


class TestWhatsNewAll:
    @pytest.mark.asyncio
    async def test_returns_all_sections(self, full_permissions):
        client = AsyncMock()
        client.get = AsyncMock(side_effect=_mock_get)
        _, result = await compose_whats_new(
            client=client,
            permissions=full_permissions,
            since="2026-05-14T00:00:00",
            scope=None,
            timezone="Europe/Berlin",
        )
        assert "Calendar" in result
        assert "Mail" in result
        assert "Teams" in result

    @pytest.mark.asyncio
    async def test_scope_mail_only_skips_calendar(self, full_permissions):
        client = AsyncMock()
        client.get = AsyncMock(side_effect=_mock_get)
        _, result = await compose_whats_new(
            client=client,
            permissions=full_permissions,
            since="2026-05-14T00:00:00",
            scope="mail",
            timezone="Europe/Berlin",
        )
        assert "Mail" in result
        assert "Calendar" not in result

    @pytest.mark.asyncio
    async def test_mail_failure_graceful(self, full_permissions):
        client = AsyncMock()

        async def _get(endpoint, params=None, headers=None):
            if "messages" in endpoint:
                raise GraphAPIError(500, "InternalError", "down")
            return await _mock_get(endpoint)

        client.get = AsyncMock(side_effect=_get)
        _, result = await compose_whats_new(
            client=client,
            permissions=full_permissions,
            since="2026-05-14T00:00:00",
            scope=None,
            timezone="Europe/Berlin",
        )
        assert "⚠️" in result
        assert "Calendar" in result

    @pytest.mark.asyncio
    async def test_no_mail_permission_shows_message(self):
        permissions = PermissionRegistry(["Calendars.ReadWrite"])
        client = AsyncMock()
        client.get = AsyncMock(return_value={"value": []})
        _, result = await compose_whats_new(
            client=client,
            permissions=permissions,
            since="2026-05-14T00:00:00",
            scope=None,
            timezone="Europe/Berlin",
        )
        assert "Mail.Read" in result


class TestWhatsNewTeamsPermalink:
    @pytest.mark.asyncio
    async def test_teams_section_includes_chat_web_url(self, full_permissions):
        client = AsyncMock()

        async def fake_get(endpoint, params=None, headers=None):
            if "/me/calendarView" in endpoint:
                return {"value": []}
            if "/me/messages" in endpoint:
                return {"value": []}
            if "/me/chats" in endpoint:
                return {"value": [{
                    "id": "19:abc@thread.v2",
                    "webUrl": "https://teams.microsoft.com/l/chat/19:abc@thread.v2/conversations",
                    "lastMessagePreview": {
                        "from": {"user": {"displayName": "Alice"}},
                        "body": {"content": "Hello there"},
                    },
                }]}
            return {}

        client.get = AsyncMock(side_effect=fake_get)

        _, result = await compose_whats_new(
            client=client,
            permissions=full_permissions,
            since="2026-05-14T00:00:00",
            scope="teams",
            timezone="Europe/Berlin",
        )

        assert "[open chat]" in result
        assert "19:abc@thread.v2" in result


async def _mock_get(endpoint, params=None, headers=None):
    if "calendarView" in endpoint:
        return {"value": [
            {
                "subject": "New meeting",
                "start": {"dateTime": "2026-05-15T09:00:00"},
                "end": {"dateTime": "2026-05-15T09:30:00"},
                "location": {"displayName": ""},
                "isOnlineMeeting": True,
                "attendees": [],
                "organizer": {"emailAddress": {"name": "Alice"}},
            }
        ]}
    if "messages" in endpoint:
        return {"value": [
            {
                "subject": "Budget update",
                "from": {"emailAddress": {"name": "Finance", "address": "finance@example.com"}},
                "receivedDateTime": "2026-05-15T10:00:00Z",
                "importance": "normal",
            }
        ]}
    if "chats" in endpoint:
        return {"value": []}
    return {"value": []}
