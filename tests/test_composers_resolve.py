"""Tests for resolve composer — all 6 non-email URL types plus email and error paths."""

import urllib.parse
from unittest.mock import AsyncMock, patch

import pytest

from ms365_intent_mcp.composers.resolve import compose_resolve
from ms365_intent_mcp.graph import GraphAPIError
from ms365_intent_mcp.permissions import PermissionRegistry
from ms365_intent_mcp.resolver import ResolvedUrl, UrlParseError


@pytest.fixture
def full_permissions():
    return PermissionRegistry([
        "ChannelMessage.Read.All",
        "Chat.ReadWrite",
        "Calendars.Read",
        "Mail.Read",
        "Sites.Read.All",
        "Files.Read",
    ])


# ---------------------------------------------------------------------------
# Email (existing tests, kept intact)
# ---------------------------------------------------------------------------

class TestResolveEmail:
    @pytest.mark.asyncio
    async def test_resolves_email_url(self, full_permissions):
        client = AsyncMock()
        client.get = AsyncMock(return_value={
            "subject": "Budget Review",
            "from": {"emailAddress": {"name": "Alice", "address": "alice@sap.com"}},
            "receivedDateTime": "2026-05-15T08:00:00Z",
            "bodyPreview": "Please review attached.",
        })

        with patch("ms365_intent_mcp.composers.resolve.resolve_url") as mock_resolve:
            mock_resolve.return_value = ResolvedUrl(
                url_type="email",
                graph_endpoint="/me/messages/AAkALg123",
                required_scope="Mail.Read",
            )
            result = await compose_resolve(
                client=client,
                permissions=full_permissions,
                url="https://outlook.office365.com/mail/id/AAkALg123",
            )
        assert "Budget Review" in result
        assert "Alice" in result

    @pytest.mark.asyncio
    async def test_unrecognised_url_returns_error(self, full_permissions):
        client = AsyncMock()

        with patch("ms365_intent_mcp.composers.resolve.resolve_url") as mock_resolve:
            mock_resolve.side_effect = UrlParseError("unknown URL")
            result = await compose_resolve(
                client=client,
                permissions=full_permissions,
                url="https://www.google.com",
            )
        assert "Unrecognised" in result

    @pytest.mark.asyncio
    async def test_missing_scope_returns_message(self):
        client = AsyncMock()
        permissions = PermissionRegistry([])

        with patch("ms365_intent_mcp.composers.resolve.resolve_url") as mock_resolve:
            mock_resolve.return_value = ResolvedUrl(
                url_type="email",
                graph_endpoint="/me/messages/123",
                required_scope="Mail.Read",
            )
            result = await compose_resolve(
                client=client,
                permissions=permissions,
                url="https://outlook.office365.com/mail/id/123",
            )
        assert "Mail.Read" in result

    @pytest.mark.asyncio
    async def test_graph_error_returns_section_error(self, full_permissions):
        client = AsyncMock()
        client.get = AsyncMock(side_effect=GraphAPIError(404, "NotFound", "message gone"))

        with patch("ms365_intent_mcp.composers.resolve.resolve_url") as mock_resolve:
            mock_resolve.return_value = ResolvedUrl(
                url_type="email",
                graph_endpoint="/me/messages/missing",
                required_scope="Mail.Read",
            )
            result = await compose_resolve(
                client=client,
                permissions=full_permissions,
                url="https://outlook.office365.com/mail/id/missing",
            )
        assert "message gone" in result


# ---------------------------------------------------------------------------
# channel_message
# ---------------------------------------------------------------------------

class TestChannelMessageResolve:
    @pytest.mark.asyncio
    async def test_channel_message_resolve(self, full_permissions):
        import json
        ctx = urllib.parse.quote(json.dumps({"groupId": "team-uuid-123", "tid": "tenant"}))
        url = f"https://teams.microsoft.com/l/message/19:chan@thread.tacv2/1234567890.123456?context={ctx}"

        client = AsyncMock()
        client.get = AsyncMock(return_value={
            "body": {"content": "Hello"},
            "from": {"user": {"displayName": "Alice"}},
            "createdDateTime": "2026-01-01T10:00:00Z",
        })

        result = await compose_resolve(client=client, permissions=full_permissions, url=url)
        assert "Teams Message" in result
        assert "Alice" in result


# ---------------------------------------------------------------------------
# chat_message
# ---------------------------------------------------------------------------

class TestChatMessageResolve:
    @pytest.mark.asyncio
    async def test_chat_message_resolve(self, full_permissions):
        url = "https://teams.microsoft.com/l/message/19:somechat@unq.gbl.spaces/1234567890.123456"

        client = AsyncMock()
        client.get = AsyncMock(return_value={
            "body": {"content": "Hey team"},
            "from": {"user": {"displayName": "Bob"}},
            "createdDateTime": "2026-01-01T09:00:00Z",
        })

        result = await compose_resolve(client=client, permissions=full_permissions, url=url)
        assert "Teams Message" in result


# ---------------------------------------------------------------------------
# meeting
# ---------------------------------------------------------------------------

class TestMeetingResolve:
    def _meeting_url(self) -> tuple[str, str]:
        """Return (url, thread_id) for a meetup-join link."""
        thread_id = "19:meeting_abc123@thread.v2"
        url = f"https://teams.microsoft.com/l/meetup-join/{urllib.parse.quote(thread_id)}/0"
        return url, thread_id

    @pytest.mark.asyncio
    async def test_meeting_found(self, full_permissions):
        url, thread_id = self._meeting_url()

        client = AsyncMock()
        client.get = AsyncMock(return_value={
            "value": [{
                "subject": "Standup",
                "start": {"dateTime": "2026-01-01T09:00:00Z", "timeZone": "UTC"},
                "end": {"dateTime": "2026-01-01T09:30:00Z", "timeZone": "UTC"},
                "organizer": {"emailAddress": {"name": "Avi"}},
                "attendees": [],
                "body": {"content": ""},
                "location": {"displayName": ""},
                "isOnlineMeeting": True,
                "onlineMeeting": {"joinUrl": f"https://teams.microsoft.com/l/meetup-join/{thread_id}/0"},
            }],
        })

        result = await compose_resolve(client=client, permissions=full_permissions, url=url)
        assert "Standup" in result

    @pytest.mark.asyncio
    async def test_meeting_not_found(self, full_permissions):
        url, _ = self._meeting_url()

        client = AsyncMock()
        client.get = AsyncMock(return_value={"value": []})

        result = await compose_resolve(client=client, permissions=full_permissions, url=url)
        assert "No matching" in result or "unavailable" in result


# ---------------------------------------------------------------------------
# onedrive_file
# ---------------------------------------------------------------------------

class TestOneDriveFileResolve:
    @pytest.mark.asyncio
    async def test_onedrive_file_resolve(self, full_permissions):
        url = "https://sap-my.sharepoint.com/personal/user_sap_com/Documents/report.xlsx"

        client = AsyncMock()
        client.get = AsyncMock(return_value={
            "name": "report.xlsx",
            "size": 1024,
            "webUrl": "https://sap-my.sharepoint.com/personal/user_sap_com/Documents/report.xlsx",
        })

        result = await compose_resolve(client=client, permissions=full_permissions, url=url)
        assert "report.xlsx" in result


# ---------------------------------------------------------------------------
# onedrive_share_link
# ---------------------------------------------------------------------------

class TestOneDriveShareLinkResolve:
    @pytest.mark.asyncio
    async def test_onedrive_share_link_resolve(self, full_permissions):
        url = "https://sap-my.sharepoint.com/:x:/r/personal/user_sap_com/_layouts/15/Doc.aspx?sourcedoc=%7Babc%7D"

        client = AsyncMock()
        client.get = AsyncMock(return_value={
            "name": "Roadmap.xlsx",
            "size": 2048,
            "webUrl": "https://sap-my.sharepoint.com/personal/user_sap_com/Shared%20Documents/Roadmap.xlsx",
        })

        result = await compose_resolve(client=client, permissions=full_permissions, url=url)
        assert "Roadmap.xlsx" in result


# ---------------------------------------------------------------------------
# sharepoint_page
# ---------------------------------------------------------------------------

class TestSharePointPageResolve:
    @pytest.mark.asyncio
    async def test_sharepoint_page_found(self, full_permissions):
        url = "https://sap.sharepoint.com/sites/MyProject/SitePages/Overview.aspx"

        site_response = {"id": "site-id-123", "displayName": "My Project", "webUrl": "https://sap.sharepoint.com/sites/MyProject"}
        lists_response = {"value": [{"id": "list-id-456"}]}
        items_response = {"value": [{
            "id": "42",
            "webUrl": "https://sap.sharepoint.com/sites/MyProject/SitePages/Overview.aspx",
            "fields": {
                "FileLeafRef": "Overview.aspx",
                "Title": "Project Overview",
                "Modified": "2026-01-01T00:00:00Z",
            },
        }]}

        client = AsyncMock()
        client.get = AsyncMock(side_effect=[site_response, lists_response, items_response])

        result = await compose_resolve(client=client, permissions=full_permissions, url=url)
        assert "SharePoint Page" in result
        assert "Project Overview" in result

    @pytest.mark.asyncio
    async def test_sharepoint_page_fallback(self, full_permissions):
        url = "https://sap.sharepoint.com/sites/MyProject/SitePages/Overview.aspx"

        site_response = {"id": "site-id-123", "displayName": "My Project", "webUrl": "https://sap.sharepoint.com/sites/MyProject"}

        client = AsyncMock()
        client.get = AsyncMock(side_effect=[
            site_response,
            GraphAPIError(404, "NotFound", "not found"),
        ])

        result = await compose_resolve(client=client, permissions=full_permissions, url=url)
        assert "SharePoint Site" in result


# ---------------------------------------------------------------------------
# _get_event_by_id helper
# ---------------------------------------------------------------------------

class TestGetEventByIdHelper:
    @pytest.mark.asyncio
    async def test_returns_event_on_success(self):
        from ms365_intent_mcp.composers.resolve import _get_event_by_id

        client = AsyncMock()
        client.get = AsyncMock(return_value={
            "subject": "Project sync",
            "start": {"dateTime": "2026-05-26T10:00:00"},
            "end": {"dateTime": "2026-05-26T10:30:00"},
        })
        event = await _get_event_by_id(client, "AAMkAGI2event-id")
        assert event["subject"] == "Project sync"
        client.get.assert_called_once()
        called_endpoint = client.get.call_args[0][0]
        assert called_endpoint == "/me/events/AAMkAGI2event-id"

    @pytest.mark.asyncio
    async def test_returns_none_on_404(self):
        from ms365_intent_mcp.composers.resolve import _get_event_by_id

        client = AsyncMock()
        client.get = AsyncMock(side_effect=GraphAPIError(404, "NotFound", "gone"))
        event = await _get_event_by_id(client, "missing-id")
        assert event is None

    @pytest.mark.asyncio
    async def test_returns_none_on_empty_id(self):
        from ms365_intent_mcp.composers.resolve import _get_event_by_id

        client = AsyncMock()
        client.get = AsyncMock()
        event = await _get_event_by_id(client, "")
        assert event is None
        client.get.assert_not_called()


# ---------------------------------------------------------------------------
# chat_thread
# ---------------------------------------------------------------------------

class TestResolveChatThread:
    @pytest.fixture
    def chat_meta(self):
        return {
            "id": "19:abc@thread.v2",
            "topic": "Project Sync",
            "chatType": "meeting",
            "webUrl": "https://teams.microsoft.com/l/chat/19:abc@thread.v2/conversations",
            "onlineMeetingInfo": {
                "calendarEventId": "AAMkevent123",
                "joinWebUrl": "https://teams.microsoft.com/l/meetup-join/19:abc@thread.v2/0",
            },
            "members": [
                {"displayName": "Alice"},
                {"displayName": "Bob"},
            ],
        }

    @pytest.fixture
    def messages_payload(self):
        return {
            "value": [
                {"from": {"user": {"displayName": "Alice"}},
                 "body": {"content": "Hello"},
                 "createdDateTime": "2026-05-26T10:00:00Z"},
                {"from": {"user": {"displayName": "Bob"}},
                 "body": {"content": "Hi"},
                 "createdDateTime": "2026-05-26T10:05:00Z"},
            ]
        }

    @pytest.fixture
    def event_payload(self):
        return {
            "subject": "Project Sync",
            "start": {"dateTime": "2026-05-26T10:00:00"},
            "end": {"dateTime": "2026-05-26T10:30:00"},
            "organizer": {"emailAddress": {"name": "Alice"}},
        }

    @pytest.mark.asyncio
    async def test_happy_path_uses_calendar_event_id(
        self, full_permissions, chat_meta, messages_payload, event_payload
    ):
        client = AsyncMock()

        async def fake_get(endpoint, params=None, headers=None):
            if endpoint == "/chats/19:abc@thread.v2":
                return chat_meta
            if endpoint == "/chats/19:abc@thread.v2/messages":
                return messages_payload
            if endpoint == "/me/events/AAMkevent123":
                return event_payload
            raise AssertionError(f"Unexpected endpoint: {endpoint}")

        client.get = AsyncMock(side_effect=fake_get)

        with patch("ms365_intent_mcp.composers.resolve.resolve_url") as mock_resolve:
            mock_resolve.return_value = ResolvedUrl(
                url_type="chat_thread",
                graph_endpoint="/chats/19:abc@thread.v2",
                required_scope="Chat.ReadWrite",
                extra={"chat_id": "19:abc@thread.v2"},
            )
            result = await compose_resolve(
                client=client,
                permissions=full_permissions,
                url="https://teams.microsoft.com/l/chat/19:abc@thread.v2/conversations",
            )

        assert "Project Sync" in result
        assert "Alice" in result
        assert "Hello" in result

    @pytest.mark.asyncio
    async def test_falls_back_to_join_web_url_when_calendar_event_id_absent(
        self, full_permissions, messages_payload, event_payload
    ):
        chat_meta = {
            "id": "19:abc@thread.v2",
            "topic": "Project Sync",
            "chatType": "meeting",
            "webUrl": "https://teams.microsoft.com/l/chat/19:abc@thread.v2/conversations",
            "onlineMeetingInfo": {
                "calendarEventId": None,
                "joinWebUrl": "https://teams.microsoft.com/l/meetup-join/19:meet@thread.v2/0",
            },
            "members": [],
        }
        client = AsyncMock()

        async def fake_get(endpoint, params=None, headers=None):
            if endpoint == "/chats/19:abc@thread.v2":
                return chat_meta
            if endpoint == "/chats/19:abc@thread.v2/messages":
                return messages_payload
            if endpoint == "/me/calendarView":
                return {"value": [{**event_payload,
                                   "onlineMeeting": {"joinUrl": "https://teams.microsoft.com/l/meetup-join/19:meet@thread.v2/0"}}]}
            raise AssertionError(f"Unexpected endpoint: {endpoint}")

        client.get = AsyncMock(side_effect=fake_get)

        with patch("ms365_intent_mcp.composers.resolve.resolve_url") as mock_resolve:
            mock_resolve.return_value = ResolvedUrl(
                url_type="chat_thread",
                graph_endpoint="/chats/19:abc@thread.v2",
                required_scope="Chat.ReadWrite",
                extra={"chat_id": "19:abc@thread.v2"},
            )
            result = await compose_resolve(
                client=client,
                permissions=full_permissions,
                url="https://teams.microsoft.com/l/chat/19:abc@thread.v2/conversations",
            )

        assert "Project Sync" in result

    @pytest.mark.asyncio
    async def test_one_to_one_chat_renders_without_meeting_block(
        self, full_permissions, messages_payload
    ):
        chat_meta = {
            "id": "19:dm@unq.gbl.spaces",
            "topic": None,
            "chatType": "oneOnOne",
            "webUrl": "https://teams.microsoft.com/l/chat/19:dm@unq.gbl.spaces/conversations",
            "onlineMeetingInfo": None,
            "members": [{"displayName": "Avi"}, {"displayName": "Alice"}],
        }
        client = AsyncMock()

        async def fake_get(endpoint, params=None, headers=None):
            if endpoint == "/chats/19:dm@unq.gbl.spaces":
                return chat_meta
            if endpoint == "/chats/19:dm@unq.gbl.spaces/messages":
                return messages_payload
            raise AssertionError(f"Unexpected endpoint: {endpoint}")

        client.get = AsyncMock(side_effect=fake_get)

        with patch("ms365_intent_mcp.composers.resolve.resolve_url") as mock_resolve:
            mock_resolve.return_value = ResolvedUrl(
                url_type="chat_thread",
                graph_endpoint="/chats/19:dm@unq.gbl.spaces",
                required_scope="Chat.ReadWrite",
                extra={"chat_id": "19:dm@unq.gbl.spaces"},
            )
            result = await compose_resolve(
                client=client,
                permissions=full_permissions,
                url="https://teams.microsoft.com/l/chat/19:dm@unq.gbl.spaces/conversations",
            )

        assert "Hello" in result
        assert "Meeting context" not in result

    @pytest.mark.asyncio
    async def test_empty_online_meeting_info_renders_without_meeting_block(
        self, full_permissions, messages_payload
    ):
        chat_meta = {
            "id": "19:abc@thread.v2",
            "topic": "Group chat",
            "chatType": "group",
            "webUrl": "https://teams.microsoft.com/l/chat/19:abc@thread.v2/conversations",
            "onlineMeetingInfo": {},
            "members": [{"displayName": "A"}, {"displayName": "B"}, {"displayName": "C"}],
        }
        client = AsyncMock()

        async def fake_get(endpoint, params=None, headers=None):
            if endpoint == "/chats/19:abc@thread.v2":
                return chat_meta
            if endpoint == "/chats/19:abc@thread.v2/messages":
                return messages_payload
            raise AssertionError(f"Unexpected endpoint: {endpoint}")

        client.get = AsyncMock(side_effect=fake_get)

        with patch("ms365_intent_mcp.composers.resolve.resolve_url") as mock_resolve:
            mock_resolve.return_value = ResolvedUrl(
                url_type="chat_thread",
                graph_endpoint="/chats/19:abc@thread.v2",
                required_scope="Chat.ReadWrite",
                extra={"chat_id": "19:abc@thread.v2"},
            )
            result = await compose_resolve(
                client=client,
                permissions=full_permissions,
                url="https://teams.microsoft.com/l/chat/19:abc@thread.v2/conversations",
            )

        assert "Group chat" in result
        assert "Meeting context" not in result

    @pytest.mark.asyncio
    async def test_messages_failure_partial_success(
        self, full_permissions, chat_meta, event_payload
    ):
        client = AsyncMock()

        async def fake_get(endpoint, params=None, headers=None):
            if endpoint == "/chats/19:abc@thread.v2":
                return chat_meta
            if endpoint == "/chats/19:abc@thread.v2/messages":
                raise GraphAPIError(500, "InternalError", "down")
            if endpoint == "/me/events/AAMkevent123":
                return event_payload
            raise AssertionError(f"Unexpected endpoint: {endpoint}")

        client.get = AsyncMock(side_effect=fake_get)

        with patch("ms365_intent_mcp.composers.resolve.resolve_url") as mock_resolve:
            mock_resolve.return_value = ResolvedUrl(
                url_type="chat_thread",
                graph_endpoint="/chats/19:abc@thread.v2",
                required_scope="Chat.ReadWrite",
                extra={"chat_id": "19:abc@thread.v2"},
            )
            result = await compose_resolve(
                client=client,
                permissions=full_permissions,
                url="https://teams.microsoft.com/l/chat/19:abc@thread.v2/conversations",
            )

        assert "Project Sync" in result
        assert "⚠️" in result or "unavailable" in result.lower()
        assert "Hello" not in result  # messages failed; no message content should render

    @pytest.mark.asyncio
    async def test_chat_failure_partial_success(
        self, full_permissions, messages_payload
    ):
        client = AsyncMock()

        async def fake_get(endpoint, params=None, headers=None):
            if endpoint == "/chats/19:abc@thread.v2":
                raise GraphAPIError(403, "Forbidden", "no access")
            if endpoint == "/chats/19:abc@thread.v2/messages":
                return messages_payload
            raise AssertionError(f"Unexpected endpoint: {endpoint}")

        client.get = AsyncMock(side_effect=fake_get)

        with patch("ms365_intent_mcp.composers.resolve.resolve_url") as mock_resolve:
            mock_resolve.return_value = ResolvedUrl(
                url_type="chat_thread",
                graph_endpoint="/chats/19:abc@thread.v2",
                required_scope="Chat.ReadWrite",
                extra={"chat_id": "19:abc@thread.v2"},
            )
            result = await compose_resolve(
                client=client,
                permissions=full_permissions,
                url="https://teams.microsoft.com/l/chat/19:abc@thread.v2/conversations",
            )

        assert "⚠️" in result or "unavailable" in result.lower()
        assert "Hello" in result  # messages still rendered

    @pytest.mark.asyncio
    async def test_calendar_fuse_404_silently_skipped(
        self, full_permissions, chat_meta, messages_payload
    ):
        client = AsyncMock()

        async def fake_get(endpoint, params=None, headers=None):
            if endpoint == "/chats/19:abc@thread.v2":
                return chat_meta
            if endpoint == "/chats/19:abc@thread.v2/messages":
                return messages_payload
            if endpoint == "/me/events/AAMkevent123":
                raise GraphAPIError(404, "NotFound", "deleted")
            raise AssertionError(f"Unexpected endpoint: {endpoint}")

        client.get = AsyncMock(side_effect=fake_get)

        with patch("ms365_intent_mcp.composers.resolve.resolve_url") as mock_resolve:
            mock_resolve.return_value = ResolvedUrl(
                url_type="chat_thread",
                graph_endpoint="/chats/19:abc@thread.v2",
                required_scope="Chat.ReadWrite",
                extra={"chat_id": "19:abc@thread.v2"},
            )
            result = await compose_resolve(
                client=client,
                permissions=full_permissions,
                url="https://teams.microsoft.com/l/chat/19:abc@thread.v2/conversations",
            )

        # Chat + messages still render; meeting block silently absent
        assert "Project Sync" in result  # chat topic
        assert "Hello" in result

    @pytest.mark.asyncio
    async def test_messages_sorted_client_side(
        self, full_permissions, chat_meta, event_payload
    ):
        out_of_order = {
            "value": [
                {"from": {"user": {"displayName": "Alice"}},
                 "body": {"content": "OLDEST"},
                 "createdDateTime": "2026-05-26T08:00:00Z"},
                {"from": {"user": {"displayName": "Bob"}},
                 "body": {"content": "NEWEST"},
                 "createdDateTime": "2026-05-26T10:00:00Z"},
                {"from": {"user": {"displayName": "Carol"}},
                 "body": {"content": "MIDDLE"},
                 "createdDateTime": "2026-05-26T09:00:00Z"},
            ]
        }
        client = AsyncMock()

        async def fake_get(endpoint, params=None, headers=None):
            if endpoint == "/chats/19:abc@thread.v2":
                return chat_meta
            if endpoint == "/chats/19:abc@thread.v2/messages":
                return out_of_order
            if endpoint == "/me/events/AAMkevent123":
                return event_payload
            raise AssertionError(f"Unexpected endpoint: {endpoint}")

        client.get = AsyncMock(side_effect=fake_get)

        with patch("ms365_intent_mcp.composers.resolve.resolve_url") as mock_resolve:
            mock_resolve.return_value = ResolvedUrl(
                url_type="chat_thread",
                graph_endpoint="/chats/19:abc@thread.v2",
                required_scope="Chat.ReadWrite",
                extra={"chat_id": "19:abc@thread.v2"},
            )
            result = await compose_resolve(
                client=client,
                permissions=full_permissions,
                url="https://teams.microsoft.com/l/chat/19:abc@thread.v2/conversations",
            )

        # Newest message must appear before middle and oldest
        newest_pos = result.find("NEWEST")
        middle_pos = result.find("MIDDLE")
        oldest_pos = result.find("OLDEST")
        assert 0 <= newest_pos < middle_pos < oldest_pos
