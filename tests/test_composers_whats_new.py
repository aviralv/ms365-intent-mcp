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


class TestWhatsNewMailFilterFormat:
    """Regression for issue #27: mail $filter must be RFC-3339 (single UTC marker)."""

    @pytest.mark.asyncio
    @pytest.mark.parametrize("since", [
        "2026-07-09T06:12:54Z",           # 'Z' shorthand
        "2026-07-09T06:12:54+00:00",      # explicit UTC offset
        "2026-07-09T06:12:54",            # naive, treated as UTC
        "2026-07-09T08:12:54+02:00",      # non-UTC offset — same instant
    ])
    async def test_mail_filter_is_rfc3339_utc(self, full_permissions, since):
        client = AsyncMock()
        captured = {}

        async def fake_get(endpoint, params=None, headers=None):
            if "/me/messages" in endpoint:
                captured["params"] = params
                return {"value": []}
            return {"value": []}

        client.get = AsyncMock(side_effect=fake_get)

        await compose_whats_new(
            client=client,
            permissions=full_permissions,
            since=since,
            scope="mail",
            timezone="Europe/Berlin",
        )

        filter_clause = captured["params"]["$filter"]
        assert filter_clause == "receivedDateTime ge 2026-07-09T06:12:54Z", filter_clause
        # Never both markers concatenated (the specific shape Graph rejected)
        assert "+00:00Z" not in filter_clause



    @pytest.mark.asyncio
    async def test_teams_section_includes_chat_web_url(self, full_permissions):
        client = AsyncMock()

        async def fake_get(endpoint, params=None, headers=None):
            if "/me/calendarView" in endpoint:
                return {"value": []}
            if "/me/messages" in endpoint:
                return {"value": []}
            if endpoint.endswith("/messages") and "/me/chats/" in endpoint:
                return {"value": [{
                    "id": "m1",
                    "createdDateTime": "2026-05-15T09:00:00Z",
                    "from": {"user": {"displayName": "Alice"}},
                    "body": {"content": "Hello there"},
                }]}
            if "/me/chats" in endpoint:
                return {"value": [{
                    "id": "19:abc@thread.v2",
                    "webUrl": "https://teams.microsoft.com/l/chat/19:abc@thread.v2/conversations",
                    "lastMessagePreview": {
                        "createdDateTime": "2026-05-15T09:00:00Z",
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


class TestWhatsNewTeamsWindowMessages:
    """Regression for issue #67: whats_new must return ALL in-window messages per
    chat, not just the latest-message preview. The bug masked an inbound reply
    whenever the user's own outbound was the most recent message in the chat.
    """

    @pytest.mark.asyncio
    async def test_inbound_reply_not_masked_by_later_outbound(self, full_permissions):
        client = AsyncMock()

        async def fake_get(endpoint, params=None, headers=None):
            if "/me/calendarView" in endpoint:
                return {"value": []}
            if "/me/messages" in endpoint:
                return {"value": []}
            if endpoint.endswith("/messages") and "/me/chats/" in endpoint:
                # Per-chat message fetch — newest first. Both are in-window.
                return {"value": [
                    {
                        "id": "m2",
                        "createdDateTime": "2026-08-12T11:33:00Z",
                        "from": {"user": {"displayName": "Avi"}},
                        "body": {"content": "sounds good, will do"},
                    },
                    {
                        "id": "m1",
                        "createdDateTime": "2026-08-12T11:17:00Z",
                        "from": {"user": {"displayName": "Counterpart"}},
                        "body": {"content": "can you review the doc?"},
                    },
                ]}
            if "/me/chats" in endpoint:
                return {"value": [{
                    "id": "19:xyz@thread.v2",
                    "topic": None,
                    "webUrl": "https://teams.microsoft.com/l/chat/19:xyz@thread.v2/conversations",
                    "lastMessagePreview": {
                        "createdDateTime": "2026-08-12T11:33:00Z",
                        "from": {"user": {"displayName": "Avi"}},
                        "body": {"content": "sounds good, will do"},
                    },
                }]}
            return {"value": []}

        client.get = AsyncMock(side_effect=fake_get)

        data, result = await compose_whats_new(
            client=client,
            permissions=full_permissions,
            since="2026-08-12T09:00:00Z",
            scope="teams",
            timezone="Europe/Berlin",
        )

        senders = {t["sender"] for t in data["teams"]}
        assert "Counterpart" in senders, f"inbound reply masked: {data['teams']}"
        assert "can you review the doc?" in result

    @pytest.mark.asyncio
    async def test_messages_before_since_are_excluded(self, full_permissions):
        client = AsyncMock()

        async def fake_get(endpoint, params=None, headers=None):
            if endpoint.endswith("/messages") and "/me/chats/" in endpoint:
                return {"value": [
                    {
                        "id": "m2",
                        "createdDateTime": "2026-08-12T11:00:00Z",
                        "from": {"user": {"displayName": "Counterpart"}},
                        "body": {"content": "in window"},
                    },
                    {
                        "id": "m1",
                        "createdDateTime": "2026-08-11T08:00:00Z",
                        "from": {"user": {"displayName": "Counterpart"}},
                        "body": {"content": "before window"},
                    },
                ]}
            if "/me/chats" in endpoint:
                return {"value": [{
                    "id": "19:xyz@thread.v2",
                    "topic": None,
                    "webUrl": "https://teams.microsoft.com/l/chat/19:xyz@thread.v2/conversations",
                    "lastMessagePreview": {
                        "createdDateTime": "2026-08-12T11:00:00Z",
                        "from": {"user": {"displayName": "Counterpart"}},
                        "body": {"content": "in window"},
                    },
                }]}
            return {"value": []}

        client.get = AsyncMock(side_effect=fake_get)

        data, _ = await compose_whats_new(
            client=client,
            permissions=full_permissions,
            since="2026-08-12T09:00:00Z",
            scope="teams",
            timezone="Europe/Berlin",
        )

        bodies = {t["body_preview"] for t in data["teams"]}
        assert "in window" in bodies
        assert "before window" not in bodies

    @pytest.mark.asyncio
    async def test_system_event_messages_excluded(self, full_permissions):
        """Live data (2026-08-12): Graph returns call-started/member-added events as
        messageType='systemEventMessage' with body '<systemEventMessage/>'. These are
        noise and must not surface (sender 'Unknown', body '<systemEventMessage/>')."""
        client = AsyncMock()

        async def fake_get(endpoint, params=None, headers=None):
            if endpoint.endswith("/messages") and "/me/chats/" in endpoint:
                return {"value": [
                    {
                        "id": "m2",
                        "messageType": "message",
                        "createdDateTime": "2026-08-12T11:10:00Z",
                        "from": {"user": {"displayName": "Counterpart"}},
                        "body": {"content": "a real message"},
                    },
                    {
                        "id": "m1",
                        "messageType": "systemEventMessage",
                        "createdDateTime": "2026-08-12T11:05:00Z",
                        "from": None,
                        "body": {"content": "<systemEventMessage/>"},
                    },
                ]}
            if "/me/chats" in endpoint:
                return {"value": [{
                    "id": "19:xyz@thread.v2",
                    "topic": None,
                    "webUrl": "https://teams.microsoft.com/l/chat/19:xyz@thread.v2/conversations",
                    "lastMessagePreview": {
                        "createdDateTime": "2026-08-12T11:10:00Z",
                        "from": {"user": {"displayName": "Counterpart"}},
                        "body": {"content": "a real message"},
                    },
                }]}
            return {"value": []}

        client.get = AsyncMock(side_effect=fake_get)

        data, _ = await compose_whats_new(
            client=client,
            permissions=full_permissions,
            since="2026-08-12T09:00:00Z",
            scope="teams",
            timezone="Europe/Berlin",
        )

        bodies = [t["body_preview"] for t in data["teams"]]
        assert bodies == ["a real message"], data["teams"]


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


class TestWhatsNewIsReadPropagation:
    @pytest.mark.asyncio
    async def test_is_read_true_propagates(self, full_permissions):
        """When Graph returns isRead=True, the structured data must have is_read=True."""
        client = AsyncMock()

        async def fake_get(endpoint, params=None, headers=None):
            if "calendarView" in endpoint:
                return {"value": []}
            if "messages" in endpoint:
                return {"value": [
                    {
                        "subject": "Read mail",
                        "from": {"emailAddress": {"name": "Alice", "address": "alice@example.com"}},
                        "receivedDateTime": "2026-07-01T10:00:00Z",
                        "importance": "normal",
                        "isRead": True,
                    }
                ]}
            return {"value": []}

        client.get = AsyncMock(side_effect=fake_get)

        data, _ = await compose_whats_new(
            client=client,
            permissions=full_permissions,
            since="2026-07-01T00:00:00",
            scope="mail",
            timezone="Europe/Berlin",
        )

        assert len(data["mail"]) == 1
        assert data["mail"][0]["is_read"] is True

    @pytest.mark.asyncio
    async def test_is_read_false_propagates(self, full_permissions):
        """When Graph returns isRead=False, the structured data must have is_read=False."""
        client = AsyncMock()

        async def fake_get(endpoint, params=None, headers=None):
            if "calendarView" in endpoint:
                return {"value": []}
            if "messages" in endpoint:
                return {"value": [
                    {
                        "subject": "Unread mail",
                        "from": {"emailAddress": {"name": "Bob", "address": "bob@example.com"}},
                        "receivedDateTime": "2026-07-01T10:00:00Z",
                        "importance": "normal",
                        "isRead": False,
                    }
                ]}
            return {"value": []}

        client.get = AsyncMock(side_effect=fake_get)

        data, _ = await compose_whats_new(
            client=client,
            permissions=full_permissions,
            since="2026-07-01T00:00:00",
            scope="mail",
            timezone="Europe/Berlin",
        )

        assert len(data["mail"]) == 1
        assert data["mail"][0]["is_read"] is False


class TestWhatsNewEventTimezones:
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
            return {"value": []}

        client.get = AsyncMock(side_effect=fake_get)

        data, _ = await compose_whats_new(
            client=client,
            permissions=full_permissions,
            since="2026-07-29T00:00:00",
            scope="calendar",
            timezone="UTC",
        )

        ev = data["events"][0]
        assert ev["start"] == "2026-07-29T14:00:00+00:00"
        assert ev["start_timezone"] == "UTC"
        assert ev["end_timezone"] == "UTC"

        allday = data["events"][1]
        assert allday["start"] == "2026-07-30"
        assert allday["start_timezone"] == "UTC"


@pytest.mark.asyncio
async def test_mail_items_carry_message_id_and_web_link(full_permissions):
    client = AsyncMock()

    async def _get(endpoint, params=None, headers=None):
        if "/me/messages" in endpoint:
            # assert the $select requests id + webLink
            assert "id" in params["$select"]
            assert "webLink" in params["$select"]
            return {"value": [{
                "id": "AAMkMSG1",
                "subject": "Hi",
                "from": {"emailAddress": {"name": "Bob", "address": "bob@x.com"}},
                "receivedDateTime": "2026-07-30T10:00:00Z",
                "importance": "normal",
                "webLink": "https://outlook.office365.com/owa/?ItemID=AAMkMSG1",
            }]}
        return {"value": []}

    client.get = AsyncMock(side_effect=_get)
    data, _ = await compose_whats_new(client, full_permissions, "2026-07-30T00:00:00Z", "mail", "UTC")
    assert data["mail"][0]["message_id"] == "AAMkMSG1"
    assert data["mail"][0]["web_link"] == "https://outlook.office365.com/owa/?ItemID=AAMkMSG1"
