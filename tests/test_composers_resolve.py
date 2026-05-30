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
            if endpoint.startswith("/chats/19:abc@thread.v2/messages"):
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
            if endpoint.startswith("/chats/19:abc@thread.v2/messages"):
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
            if endpoint.startswith("/chats/19:dm@unq.gbl.spaces/messages"):
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
            if endpoint.startswith("/chats/19:abc@thread.v2/messages"):
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
            if endpoint.startswith("/chats/19:abc@thread.v2/messages"):
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
            if endpoint.startswith("/chats/19:abc@thread.v2/messages"):
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
            if endpoint.startswith("/chats/19:abc@thread.v2/messages"):
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
            if endpoint.startswith("/chats/19:abc@thread.v2/messages"):
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


# ---------------------------------------------------------------------------
# _parse_iso_duration
# ---------------------------------------------------------------------------

class TestParseIsoDuration:
    def test_basic(self):
        from ms365_intent_mcp.composers.resolve import _parse_iso_duration
        assert _parse_iso_duration("PT25M38.4845646S") == "25m38s"

    def test_only_minutes(self):
        from ms365_intent_mcp.composers.resolve import _parse_iso_duration
        assert _parse_iso_duration("PT5M") == "5m"

    def test_only_seconds(self):
        from ms365_intent_mcp.composers.resolve import _parse_iso_duration
        assert _parse_iso_duration("PT30S") == "30s"

    def test_with_hours(self):
        from ms365_intent_mcp.composers.resolve import _parse_iso_duration
        assert _parse_iso_duration("PT1H30M") == "1h30m"

    def test_zero(self):
        from ms365_intent_mcp.composers.resolve import _parse_iso_duration
        assert _parse_iso_duration("PT0S") == "0s"

    def test_fractional_only_seconds(self):
        from ms365_intent_mcp.composers.resolve import _parse_iso_duration
        assert _parse_iso_duration("PT1.5S") == "1s"

    def test_empty_returns_empty(self):
        from ms365_intent_mcp.composers.resolve import _parse_iso_duration
        assert _parse_iso_duration("") == ""

    def test_malformed_returns_empty(self):
        from ms365_intent_mcp.composers.resolve import _parse_iso_duration
        assert _parse_iso_duration("not-a-duration") == ""


# ---------------------------------------------------------------------------
# _build_member_name_map
# ---------------------------------------------------------------------------

class TestBuildMemberNameMap:
    def test_basic(self):
        from ms365_intent_mcp.composers.resolve import _build_member_name_map
        chat = {
            "members": [
                {"userId": "u1", "displayName": "Alice"},
                {"userId": "u2", "displayName": "Bob"},
            ]
        }
        assert _build_member_name_map(chat) == {"u1": "Alice", "u2": "Bob"}

    def test_skips_members_without_user_id(self):
        from ms365_intent_mcp.composers.resolve import _build_member_name_map
        chat = {
            "members": [
                {"userId": "u1", "displayName": "Alice"},
                {"displayName": "Anonymous Guest"},  # no userId
            ]
        }
        assert _build_member_name_map(chat) == {"u1": "Alice"}

    def test_skips_members_without_display_name(self):
        from ms365_intent_mcp.composers.resolve import _build_member_name_map
        chat = {
            "members": [
                {"userId": "u1", "displayName": "Alice"},
                {"userId": "u2"},  # no displayName
            ]
        }
        assert _build_member_name_map(chat) == {"u1": "Alice"}

    def test_empty_members(self):
        from ms365_intent_mcp.composers.resolve import _build_member_name_map
        assert _build_member_name_map({"members": []}) == {}

    def test_missing_members_key(self):
        from ms365_intent_mcp.composers.resolve import _build_member_name_map
        assert _build_member_name_map({}) == {}

    def test_none_chat(self):
        from ms365_intent_mcp.composers.resolve import _build_member_name_map
        assert _build_member_name_map(None) == {}


# ---------------------------------------------------------------------------
# _message_entry
# ---------------------------------------------------------------------------

class TestMessageEntry:
    def test_resolves_sender_via_name_map(self):
        from ms365_intent_mcp.composers.resolve import _message_entry
        msg = {
            "createdDateTime": "2026-05-29T10:00:00Z",
            "from": {"user": {"id": "u1", "displayName": None}},
            "body": {"content": "Hello"},
        }
        entry = _message_entry(msg, {"u1": "Alice"})
        assert entry["kind"] == "message"
        assert entry["sender"] == "Alice"
        assert entry["body"] == "Hello"
        assert entry["is_body_empty"] is False
        assert entry["ts"] == "2026-05-29T10:00:00Z"

    def test_falls_back_to_graph_user_displayname(self):
        from ms365_intent_mcp.composers.resolve import _message_entry
        msg = {
            "createdDateTime": "2026-05-29T10:00:00Z",
            "from": {"user": {"id": "u-unknown", "displayName": "Bob (Graph)"}},
            "body": {"content": "Hi"},
        }
        entry = _message_entry(msg, {})  # empty map
        assert entry["sender"] == "Bob (Graph)"

    def test_falls_back_to_application_displayname(self):
        from ms365_intent_mcp.composers.resolve import _message_entry
        msg = {
            "createdDateTime": "2026-05-29T10:00:00Z",
            "from": {"application": {"displayName": "Workflow Bot"}},
            "body": {"content": "automated"},
        }
        entry = _message_entry(msg, {})
        assert entry["sender"] == "Workflow Bot"

    def test_falls_back_to_unknown(self):
        from ms365_intent_mcp.composers.resolve import _message_entry
        msg = {
            "createdDateTime": "2026-05-29T10:00:00Z",
            "from": None,
            "body": {"content": "mystery"},
        }
        entry = _message_entry(msg, {})
        assert entry["sender"] == "Unknown"

    def test_truncates_long_body_with_ellipsis(self):
        from ms365_intent_mcp.composers.resolve import _message_entry
        msg = {
            "createdDateTime": "2026-05-29T10:00:00Z",
            "from": {"user": {"id": "u1", "displayName": "A"}},
            "body": {"content": "x" * 600},
        }
        entry = _message_entry(msg, {})
        assert entry["body"].endswith("…")
        assert len(entry["body"]) == 501  # 500 chars + "…"

    def test_at_500_chars_no_ellipsis(self):
        from ms365_intent_mcp.composers.resolve import _message_entry
        msg = {
            "createdDateTime": "2026-05-29T10:00:00Z",
            "from": {"user": {"id": "u1", "displayName": "A"}},
            "body": {"content": "x" * 500},
        }
        entry = _message_entry(msg, {})
        assert entry["body"] == "x" * 500
        assert "…" not in entry["body"]

    def test_html_stripped(self):
        from ms365_intent_mcp.composers.resolve import _message_entry
        msg = {
            "createdDateTime": "2026-05-29T10:00:00Z",
            "from": {"user": {"id": "u1", "displayName": "A"}},
            "body": {"content": "<p>Hello <b>world</b></p>"},
        }
        entry = _message_entry(msg, {})
        assert entry["body"] == "Hello world"

    def test_at_mention_only_renders_inner_text(self):
        from ms365_intent_mcp.composers.resolve import _message_entry
        msg = {
            "createdDateTime": "2026-05-29T10:00:00Z",
            "from": {"user": {"id": "u1", "displayName": "A"}},
            "body": {"content": '<at id="123">@Avi</at>'},
        }
        entry = _message_entry(msg, {})
        assert entry["body"] == "@Avi"
        assert entry["is_body_empty"] is False

    def test_empty_body_marked_is_body_empty(self):
        from ms365_intent_mcp.composers.resolve import _message_entry
        msg = {
            "createdDateTime": "2026-05-29T10:00:00Z",
            "from": {"user": {"id": "u1", "displayName": "A"}},
            "body": {"content": ""},
        }
        entry = _message_entry(msg, {})
        assert entry["is_body_empty"] is True
        assert entry["body"] == ""

    def test_html_only_body_marked_is_body_empty(self):
        from ms365_intent_mcp.composers.resolve import _message_entry
        msg = {
            "createdDateTime": "2026-05-29T10:00:00Z",
            "from": {"user": {"id": "u1", "displayName": "A"}},
            "body": {"content": "<systemEventMessage/>"},
        }
        entry = _message_entry(msg, {})
        assert entry["is_body_empty"] is True


# ---------------------------------------------------------------------------
# _event_entry
# ---------------------------------------------------------------------------

class TestEventEntry:
    def test_members_added_with_names(self):
        from ms365_intent_mcp.composers.resolve import _event_entry
        msg = {
            "createdDateTime": "2026-05-29T10:00:00Z",
            "eventDetail": {
                "@odata.type": "#microsoft.graph.membersAddedEventMessageDetail",
                "members": [
                    {"displayName": "Carol Smith"},
                    {"displayName": "Dan Jones"},
                ],
            },
        }
        entry = _event_entry(msg)
        assert entry["kind"] == "event"
        assert entry["event_type"] == "membersAdded"
        assert entry["summary"] == "Member added: Carol Smith, Dan Jones"

    def test_members_added_empty_falls_back(self):
        from ms365_intent_mcp.composers.resolve import _event_entry
        msg = {
            "createdDateTime": "2026-05-29T10:00:00Z",
            "eventDetail": {
                "@odata.type": "#microsoft.graph.membersAddedEventMessageDetail",
                "members": [],
            },
        }
        entry = _event_entry(msg)
        assert entry["summary"] == "Member added: (someone)"

    def test_members_deleted(self):
        from ms365_intent_mcp.composers.resolve import _event_entry
        msg = {
            "createdDateTime": "2026-05-29T10:00:00Z",
            "eventDetail": {
                "@odata.type": "#microsoft.graph.membersDeletedEventMessageDetail",
                "members": [{"displayName": "Eve"}],
            },
        }
        entry = _event_entry(msg)
        assert entry["event_type"] == "membersDeleted"
        assert entry["summary"] == "Member removed: Eve"

    def test_chat_renamed(self):
        from ms365_intent_mcp.composers.resolve import _event_entry
        msg = {
            "createdDateTime": "2026-05-29T10:00:00Z",
            "eventDetail": {
                "@odata.type": "#microsoft.graph.chatRenamedEventMessageDetail",
                "chatDisplayName": "Project Sync 2026",
            },
        }
        entry = _event_entry(msg)
        assert entry["event_type"] == "chatRenamed"
        assert entry["summary"] == 'Renamed to "Project Sync 2026"'

    def test_chat_renamed_empty(self):
        from ms365_intent_mcp.composers.resolve import _event_entry
        msg = {
            "createdDateTime": "2026-05-29T10:00:00Z",
            "eventDetail": {
                "@odata.type": "#microsoft.graph.chatRenamedEventMessageDetail",
                "chatDisplayName": "",
            },
        }
        entry = _event_entry(msg)
        assert entry["summary"] == "Chat renamed"

    def test_unknown_type(self):
        from ms365_intent_mcp.composers.resolve import _event_entry
        msg = {
            "createdDateTime": "2026-05-29T10:00:00Z",
            "eventDetail": {
                "@odata.type": "#microsoft.graph.tabUpdatedEventMessageDetail",
            },
        }
        entry = _event_entry(msg)
        assert entry["event_type"] == "unknown"
        assert entry["summary"] == "system event"

    def test_members_added_resolves_ids_via_name_map(self):
        """Bug 2: Graph returns id-only members with displayName=null.
        Use name_map to resolve via the chat's member list."""
        from ms365_intent_mcp.composers.resolve import _event_entry
        msg = {
            "createdDateTime": "2026-05-29T10:00:00Z",
            "eventDetail": {
                "@odata.type": "#microsoft.graph.membersAddedEventMessageDetail",
                "members": [
                    {"id": "u1", "displayName": None, "userIdentityType": "aadUser"},
                    {"id": "u2", "displayName": None, "userIdentityType": "aadUser"},
                ],
            },
        }
        entry = _event_entry(msg, {"u1": "Alice", "u2": "Bob"})
        assert entry["summary"] == "Member added: Alice, Bob"

    def test_members_added_falls_back_to_displayname_when_id_missing_from_map(self):
        from ms365_intent_mcp.composers.resolve import _event_entry
        msg = {
            "createdDateTime": "2026-05-29T10:00:00Z",
            "eventDetail": {
                "@odata.type": "#microsoft.graph.membersAddedEventMessageDetail",
                "members": [
                    {"id": "u-unknown", "displayName": "Carol"},
                ],
            },
        }
        entry = _event_entry(msg, {})
        assert entry["summary"] == "Member added: Carol"

    def test_members_added_someone_when_ids_unresolvable(self):
        from ms365_intent_mcp.composers.resolve import _event_entry
        msg = {
            "createdDateTime": "2026-05-29T10:00:00Z",
            "eventDetail": {
                "@odata.type": "#microsoft.graph.membersAddedEventMessageDetail",
                "members": [
                    {"id": "u-unknown", "displayName": None},
                ],
            },
        }
        entry = _event_entry(msg, {})
        assert entry["summary"] == "Member added: (someone)"


# ---------------------------------------------------------------------------
# _group_call_events
# ---------------------------------------------------------------------------

class TestGroupCallEvents:
    def _recording_event(self, call_id, ts, status, url="", duration=""):
        return {
            "createdDateTime": ts,
            "eventDetail": {
                "@odata.type": "#microsoft.graph.callRecordingEventMessageDetail",
                "callId": call_id,
                "callRecordingStatus": status,
                "callRecordingUrl": url,
                "callRecordingDuration": duration,
                "initiator": {"user": {"id": "u1", "displayName": None}},
            },
        }

    def _transcript_event(self, call_id, ts):
        return {
            "createdDateTime": ts,
            "eventDetail": {
                "@odata.type": "#microsoft.graph.callTranscriptEventMessageDetail",
                "callId": call_id,
            },
        }

    def test_single_call(self):
        from ms365_intent_mcp.composers.resolve import _group_call_events
        events = [
            self._recording_event("c1", "2026-05-29T10:00:00Z", "initial"),
            self._recording_event("c1", "2026-05-29T10:05:00Z", "chunkFinished"),
            self._recording_event("c1", "2026-05-29T10:25:00Z", "success",
                                  url="https://recording.example/c1.mp4",
                                  duration="PT25M0S"),
        ]
        out = _group_call_events(events, {"u1": "Alice"})
        assert len(out) == 1
        call = out[0]
        assert call["kind"] == "call"
        assert call["ts"] == "2026-05-29T10:00:00Z"
        assert call["end_ts"] == "2026-05-29T10:25:00Z"
        assert call["duration"] == "25m0s"
        assert call["recording_url"] == "https://recording.example/c1.mp4"
        assert call["initiator"] == "Alice"
        assert call["transcript_ready"] is False

    def test_multiple_calls_separate_entries(self):
        from ms365_intent_mcp.composers.resolve import _group_call_events
        events = [
            self._recording_event("c1", "2026-05-29T10:00:00Z", "success",
                                  url="https://r1", duration="PT5M0S"),
            self._recording_event("c2", "2026-05-29T11:00:00Z", "success",
                                  url="https://r2", duration="PT10M0S"),
        ]
        out = _group_call_events(events, {})
        assert len(out) == 2
        urls = {c["recording_url"] for c in out}
        assert urls == {"https://r1", "https://r2"}

    def test_recording_url_prefers_success_status(self):
        from ms365_intent_mcp.composers.resolve import _group_call_events
        events = [
            self._recording_event("c1", "2026-05-29T10:00:00Z", "chunkFinished",
                                  url="https://chunk.example/temp.mp4"),
            self._recording_event("c1", "2026-05-29T10:25:00Z", "success",
                                  url="https://final.example/c1.mp4"),
        ]
        out = _group_call_events(events, {})
        assert out[0]["recording_url"] == "https://final.example/c1.mp4"

    def test_recording_url_pending_no_success(self):
        from ms365_intent_mcp.composers.resolve import _group_call_events
        events = [
            self._recording_event("c1", "2026-05-29T10:00:00Z", "initial"),
            self._recording_event("c1", "2026-05-29T10:05:00Z", "chunkFinished",
                                  url="https://chunk.example/temp.mp4"),
        ]
        out = _group_call_events(events, {})
        assert out[0]["recording_url"] == ""

    def test_status_case_insensitive(self):
        from ms365_intent_mcp.composers.resolve import _group_call_events
        events = [
            self._recording_event("c1", "2026-05-29T10:00:00Z", "Success",
                                  url="https://final.example/c1.mp4"),
        ]
        out = _group_call_events(events, {})
        assert out[0]["recording_url"] == "https://final.example/c1.mp4"

    def test_transcript_ready_flag(self):
        from ms365_intent_mcp.composers.resolve import _group_call_events
        events = [
            self._recording_event("c1", "2026-05-29T10:00:00Z", "success",
                                  url="https://r1"),
            self._transcript_event("c1", "2026-05-29T10:30:00Z"),
        ]
        out = _group_call_events(events, {})
        assert out[0]["transcript_ready"] is True

    def test_duration_from_call_ended_event(self):
        from ms365_intent_mcp.composers.resolve import _group_call_events
        events = [
            {
                "createdDateTime": "2026-05-29T10:30:00Z",
                "eventDetail": {
                    "@odata.type": "#microsoft.graph.callEndedEventMessageDetail",
                    "callId": "c1",
                    "callDuration": "PT45M0S",
                },
            },
        ]
        out = _group_call_events(events, {})
        assert out[0]["duration"] == "45m0s"

    def test_initiator_resolved_via_map(self):
        from ms365_intent_mcp.composers.resolve import _group_call_events
        events = [self._recording_event("c1", "2026-05-29T10:00:00Z", "success",
                                        url="https://r1")]
        out = _group_call_events(events, {"u1": "Alice"})
        assert out[0]["initiator"] == "Alice"

    def test_initiator_falls_back_to_graph_displayname(self):
        from ms365_intent_mcp.composers.resolve import _group_call_events
        event = {
            "createdDateTime": "2026-05-29T10:00:00Z",
            "eventDetail": {
                "@odata.type": "#microsoft.graph.callRecordingEventMessageDetail",
                "callId": "c1",
                "callRecordingStatus": "success",
                "callRecordingUrl": "https://r1",
                "initiator": {"user": {"id": "u-unknown", "displayName": "Bob (Graph)"}},
            },
        }
        out = _group_call_events([event], {})
        assert out[0]["initiator"] == "Bob (Graph)"

    def test_initiator_none_when_unresolvable(self):
        from ms365_intent_mcp.composers.resolve import _group_call_events
        event = {
            "createdDateTime": "2026-05-29T10:00:00Z",
            "eventDetail": {
                "@odata.type": "#microsoft.graph.callRecordingEventMessageDetail",
                "callId": "c1",
                "callRecordingStatus": "success",
                "callRecordingUrl": "https://r1",
                "initiator": None,
            },
        }
        out = _group_call_events([event], {})
        assert out[0]["initiator"] is None

    def test_call_unknown_when_no_call_id(self):
        from ms365_intent_mcp.composers.resolve import _group_call_events
        event = {
            "createdDateTime": "2026-05-29T10:00:00Z",
            "eventDetail": {
                "@odata.type": "#microsoft.graph.callRecordingEventMessageDetail",
                "callRecordingStatus": "success",
            },
        }
        out = _group_call_events([event], {})
        assert len(out) == 1
        assert out[0]["kind"] == "event"
        assert out[0]["event_type"] == "call_unknown"
        assert out[0]["summary"] == "Call event (no callId)"

    def test_duration_prefers_success_event_over_initial_pt0s(self):
        """Bug 1 regression: PT0S on initial event must not lock in duration."""
        from ms365_intent_mcp.composers.resolve import _group_call_events
        events = [
            self._recording_event("c1", "2026-05-27T11:46:27Z", "initial", duration="PT0S"),
            self._recording_event("c1", "2026-05-27T12:12:05Z", "chunkFinished",
                                  duration="PT25M38S"),
            self._recording_event("c1", "2026-05-27T12:12:26Z", "success",
                                  url="https://r1", duration="PT25M38S"),
        ]
        out = _group_call_events(events, {})
        assert out[0]["duration"] == "25m38s"

    def test_duration_falls_back_to_chunk_when_no_success(self):
        """Recording in progress: success event hasn't fired yet. Use the
        latest non-zero chunk duration instead of the initial PT0S."""
        from ms365_intent_mcp.composers.resolve import _group_call_events
        events = [
            self._recording_event("c1", "2026-05-27T11:46:27Z", "initial", duration="PT0S"),
            self._recording_event("c1", "2026-05-27T12:12:05Z", "chunkFinished",
                                  duration="PT25M38S"),
        ]
        out = _group_call_events(events, {})
        assert out[0]["duration"] == "25m38s"

    def test_duration_from_call_ended_overrides_recording_pt0s(self):
        """If call recording fires first with PT0S but callEnded has real duration."""
        from ms365_intent_mcp.composers.resolve import _group_call_events
        events = [
            self._recording_event("c1", "2026-05-27T10:00:00Z", "initial", duration="PT0S"),
            {
                "createdDateTime": "2026-05-27T10:30:00Z",
                "eventDetail": {
                    "@odata.type": "#microsoft.graph.callEndedEventMessageDetail",
                    "callId": "c1",
                    "callDuration": "PT30M0S",
                },
            },
        ]
        out = _group_call_events(events, {})
        assert out[0]["duration"] == "30m0s"


# ---------------------------------------------------------------------------
# _normalize_chat_entries
# ---------------------------------------------------------------------------

class TestNormalizeChatEntries:
    def test_partitions_messages_and_events(self):
        from ms365_intent_mcp.composers.resolve import _normalize_chat_entries
        messages = [
            {
                "createdDateTime": "2026-05-29T10:00:00Z",
                "from": {"user": {"id": "u1", "displayName": "Alice"}},
                "body": {"content": "Hi"},
            },
            {
                "createdDateTime": "2026-05-29T10:05:00Z",
                "from": None,
                "body": {"content": "<systemEventMessage/>"},
                "eventDetail": {
                    "@odata.type": "#microsoft.graph.callRecordingEventMessageDetail",
                    "callId": "c1",
                    "callRecordingStatus": "success",
                    "callRecordingUrl": "https://r1",
                },
            },
        ]
        entries = _normalize_chat_entries(messages, {})
        kinds = [e["kind"] for e in entries]
        assert "message" in kinds
        assert "call" in kinds

    def test_sorts_desc_by_ts(self):
        from ms365_intent_mcp.composers.resolve import _normalize_chat_entries
        messages = [
            {
                "createdDateTime": "2026-05-29T08:00:00Z",
                "from": {"user": {"id": "u1", "displayName": "A"}},
                "body": {"content": "OLDEST"},
            },
            {
                "createdDateTime": "2026-05-29T10:00:00Z",
                "from": {"user": {"id": "u1", "displayName": "A"}},
                "body": {"content": "NEWEST"},
            },
            {
                "createdDateTime": "2026-05-29T09:00:00Z",
                "from": {"user": {"id": "u1", "displayName": "A"}},
                "body": {"content": "MIDDLE"},
            },
        ]
        entries = _normalize_chat_entries(messages, {})
        assert entries[0]["body"] == "NEWEST"
        assert entries[1]["body"] == "MIDDLE"
        assert entries[2]["body"] == "OLDEST"

    def test_caps_at_25_entries(self):
        from ms365_intent_mcp.composers.resolve import _normalize_chat_entries
        messages = [
            {
                "createdDateTime": f"2026-05-29T10:{i:02d}:00Z",
                "from": {"user": {"id": "u1", "displayName": "A"}},
                "body": {"content": f"msg{i}"},
            }
            for i in range(40)
        ]
        entries = _normalize_chat_entries(messages, {})
        assert len(entries) == 25

    def test_empty_input(self):
        from ms365_intent_mcp.composers.resolve import _normalize_chat_entries
        assert _normalize_chat_entries([], {}) == []

    def test_classifies_member_event_as_event(self):
        from ms365_intent_mcp.composers.resolve import _normalize_chat_entries
        messages = [
            {
                "createdDateTime": "2026-05-29T10:00:00Z",
                "eventDetail": {
                    "@odata.type": "#microsoft.graph.membersAddedEventMessageDetail",
                    "members": [{"displayName": "Carol"}],
                },
            },
        ]
        entries = _normalize_chat_entries(messages, {})
        assert len(entries) == 1
        assert entries[0]["kind"] == "event"
        assert entries[0]["event_type"] == "membersAdded"


# ---------------------------------------------------------------------------
# _paginate_chat_messages
# ---------------------------------------------------------------------------

class TestPaginateChatMessages:
    @pytest.mark.asyncio
    async def test_single_page_no_next_link(self):
        from ms365_intent_mcp.composers.resolve import _paginate_chat_messages
        client = AsyncMock()
        client.get = AsyncMock(return_value={
            "value": [{"id": "m1"}, {"id": "m2"}],
        })
        messages, err = await _paginate_chat_messages(client, "chat1")
        assert len(messages) == 2
        assert err is None

    @pytest.mark.asyncio
    async def test_follows_next_link(self):
        from ms365_intent_mcp.composers.resolve import _paginate_chat_messages
        client = AsyncMock()
        responses = [
            {"value": [{"id": "m1"}, {"id": "m2"}],
             "@odata.nextLink": "https://graph.microsoft.com/v1.0/chats/chat1/messages?$skiptoken=p2"},
            {"value": [{"id": "m3"}, {"id": "m4"}]},
        ]
        client.get = AsyncMock(side_effect=responses)
        messages, err = await _paginate_chat_messages(client, "chat1")
        assert [m["id"] for m in messages] == ["m1", "m2", "m3", "m4"]
        assert err is None

    @pytest.mark.asyncio
    async def test_dedup_by_message_id(self):
        from ms365_intent_mcp.composers.resolve import _paginate_chat_messages
        client = AsyncMock()
        responses = [
            {"value": [{"id": "m1"}, {"id": "m2"}],
             "@odata.nextLink": "https://graph.microsoft.com/v1.0/chats/chat1/messages?$skiptoken=p2"},
            {"value": [{"id": "m2"}, {"id": "m3"}]},  # m2 duplicates
        ]
        client.get = AsyncMock(side_effect=responses)
        messages, err = await _paginate_chat_messages(client, "chat1")
        assert [m["id"] for m in messages] == ["m1", "m2", "m3"]

    @pytest.mark.asyncio
    async def test_includes_messages_without_id(self):
        from ms365_intent_mcp.composers.resolve import _paginate_chat_messages
        client = AsyncMock()
        client.get = AsyncMock(return_value={
            "value": [{"id": "m1"}, {"createdDateTime": "x"}],  # second has no id
        })
        messages, err = await _paginate_chat_messages(client, "chat1")
        assert len(messages) == 2  # both included; id-less NOT silently dropped

    @pytest.mark.asyncio
    async def test_first_page_failure_raises(self):
        from ms365_intent_mcp.composers.resolve import _paginate_chat_messages
        client = AsyncMock()
        client.get = AsyncMock(side_effect=GraphAPIError(500, "X", "down"))
        with pytest.raises(GraphAPIError):
            await _paginate_chat_messages(client, "chat1")

    @pytest.mark.asyncio
    async def test_mid_stream_failure_returns_partial(self):
        from ms365_intent_mcp.composers.resolve import _paginate_chat_messages
        client = AsyncMock()
        responses = [
            {"value": [{"id": "m1"}, {"id": "m2"}],
             "@odata.nextLink": "https://graph.microsoft.com/v1.0/chats/chat1/messages?$skiptoken=p2"},
            GraphAPIError(500, "X", "page 2 down"),
        ]
        client.get = AsyncMock(side_effect=responses)
        messages, err = await _paginate_chat_messages(client, "chat1")
        assert [m["id"] for m in messages] == ["m1", "m2"]
        assert err is not None
        assert "page 2" in err

    @pytest.mark.asyncio
    async def test_caps_at_max_messages(self):
        from ms365_intent_mcp.composers.resolve import _paginate_chat_messages
        client = AsyncMock()
        page1 = {"value": [{"id": f"m{i}"} for i in range(60)],
                 "@odata.nextLink": "https://graph.microsoft.com/v1.0/chats/chat1/messages?$skiptoken=p2"}
        page2 = {"value": [{"id": f"m{60+i}"} for i in range(60)]}
        client.get = AsyncMock(side_effect=[page1, page2])
        messages, err = await _paginate_chat_messages(client, "chat1", max_messages=100)
        assert len(messages) == 100
        assert err is None

    @pytest.mark.asyncio
    async def test_caps_at_max_pages(self):
        from ms365_intent_mcp.composers.resolve import _paginate_chat_messages
        client = AsyncMock()
        # Always returns a nextLink; only max_pages stops it
        always_paginated = {
            "value": [{"id": f"m{i}"} for i in range(2)],
            "@odata.nextLink": "https://graph.microsoft.com/v1.0/chats/chat1/messages?$skiptoken=loop",
        }
        client.get = AsyncMock(return_value=always_paginated)
        messages, err = await _paginate_chat_messages(client, "chat1", max_pages=3)
        assert client.get.call_count == 3
        assert err is None
