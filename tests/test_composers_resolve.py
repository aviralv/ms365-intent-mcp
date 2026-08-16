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
    return PermissionRegistry(
        [
            "ChannelMessage.Read.All",
            "Chat.ReadWrite",
            "Calendars.Read",
            "Mail.Read",
            "Sites.Read.All",
            "Files.Read",
        ]
    )


# ---------------------------------------------------------------------------
# Email (existing tests, kept intact)
# ---------------------------------------------------------------------------


class TestResolveEmail:
    @pytest.mark.asyncio
    async def test_resolves_email_url(self, full_permissions):
        client = AsyncMock()
        client.get = AsyncMock(
            return_value={
                "subject": "Budget Review",
                "from": {"emailAddress": {"name": "Alice", "address": "alice@example.com"}},
                "receivedDateTime": "2026-05-15T08:00:00Z",
                "bodyPreview": "Please review attached.",
            }
        )

        with patch("ms365_intent_mcp.composers.resolve.resolve_url") as mock_resolve:
            mock_resolve.return_value = ResolvedUrl(
                url_type="email",
                graph_endpoint="/me/messages/AAkALg123",
                required_scope="Mail.Read",
            )
            _, result = await compose_resolve(
                client=client,
                permissions=full_permissions,
                url="https://outlook.office365.com/mail/id/AAkALg123",
            )
        assert "Budget Review" in result
        assert "Alice" in result

    @pytest.mark.asyncio
    async def test_email_fetch_requests_plain_text_body(self, full_permissions):
        """Prefer: outlook.body-content-type=\"text\" lets Graph return
        server-generated plain text instead of HTML — cleaner for LLM
        consumption, no client-side stripping needed."""
        client = AsyncMock()
        client.get = AsyncMock(
            return_value={
                "subject": "x",
                "from": {"emailAddress": {"name": "A"}},
                "receivedDateTime": "2026-05-15T08:00:00Z",
                "body": {"contentType": "text", "content": "plain text body"},
            }
        )

        with patch("ms365_intent_mcp.composers.resolve.resolve_url") as mock_resolve:
            mock_resolve.return_value = ResolvedUrl(
                url_type="email",
                graph_endpoint="/me/messages/AAA",
                required_scope="Mail.Read",
            )
            await compose_resolve(
                client=client,
                permissions=full_permissions,
                url="https://outlook.office365.com/mail/id/AAA",
            )
        _, kwargs = client.get.call_args
        headers = kwargs.get("headers") or {}
        assert headers.get("Prefer") == 'outlook.body-content-type="text"'
        params = kwargs.get("params") or {}
        assert "body" in params.get("$select", "")

    @pytest.mark.asyncio
    async def test_email_full_body_renders_in_output(self, full_permissions):
        """End-to-end at the composer level: body.content flows through."""
        long_body = (
            "Yes exactly.\n\nDetail 1: repro on tenant X.\n\n"
            "Detail 2: repro requires empty input list.\n\n"
            "Regards, Alessia"
        )
        client = AsyncMock()
        client.get = AsyncMock(
            return_value={
                "subject": "RE: Additional Feedback on Agent",
                "from": {"emailAddress": {"name": "Alessia"}},
                "receivedDateTime": "2026-07-02T09:00:00Z",
                "body": {"contentType": "text", "content": long_body},
                "bodyPreview": long_body[:200],
            }
        )

        with patch("ms365_intent_mcp.composers.resolve.resolve_url") as mock_resolve:
            mock_resolve.return_value = ResolvedUrl(
                url_type="email",
                graph_endpoint="/me/messages/AAA",
                required_scope="Mail.Read",
            )
            _, result = await compose_resolve(
                client=client,
                permissions=full_permissions,
                url="https://outlook.office365.com/mail/id/AAA",
            )
        assert "Detail 2: repro requires empty input list" in result
        assert "Regards, Alessia" in result

    @pytest.mark.asyncio
    async def test_unrecognised_url_returns_error(self, full_permissions):
        client = AsyncMock()

        with patch("ms365_intent_mcp.composers.resolve.resolve_url") as mock_resolve:
            mock_resolve.side_effect = UrlParseError("unknown URL")
            _, result = await compose_resolve(
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
            _, result = await compose_resolve(
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
            _, result = await compose_resolve(
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
        client.get = AsyncMock(
            return_value={
                "body": {"content": "Hello"},
                "from": {"user": {"displayName": "Alice"}},
                "createdDateTime": "2026-01-01T10:00:00Z",
            }
        )

        _, result = await compose_resolve(client=client, permissions=full_permissions, url=url)
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
        client.get = AsyncMock(
            return_value={
                "body": {"content": "Hey team"},
                "from": {"user": {"displayName": "Bob"}},
                "createdDateTime": "2026-01-01T09:00:00Z",
            }
        )

        _, result = await compose_resolve(client=client, permissions=full_permissions, url=url)
        assert "Teams Message" in result

    @pytest.mark.asyncio
    async def test_chat_message_surfaces_parent_thread_url(self, full_permissions):
        """#37 Option 3: resolving a chat message link must hand back the
        parent chat thread URL so the caller can re-resolve the whole thread."""
        url = "https://teams.microsoft.com/l/message/19:somechat@unq.gbl.spaces/1234567890.123456"

        client = AsyncMock()
        client.get = AsyncMock(
            return_value={
                "body": {"content": "Hey team"},
                "from": {"user": {"displayName": "Bob"}},
                "createdDateTime": "2026-01-01T09:00:00Z",
            }
        )

        structured, _ = await compose_resolve(client=client, permissions=full_permissions, url=url)
        data = structured["data"]
        assert data["chat_id"] == "19:somechat@unq.gbl.spaces"
        assert data["chat_url"] == "https://teams.microsoft.com/l/chat/19:somechat@unq.gbl.spaces"


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
        client.get = AsyncMock(
            return_value={
                "value": [
                    {
                        "subject": "Standup",
                        "start": {"dateTime": "2026-01-01T09:00:00Z", "timeZone": "UTC"},
                        "end": {"dateTime": "2026-01-01T09:30:00Z", "timeZone": "UTC"},
                        "organizer": {"emailAddress": {"name": "Avi"}},
                        "attendees": [],
                        "body": {"content": ""},
                        "location": {"displayName": ""},
                        "isOnlineMeeting": True,
                        "onlineMeeting": {
                            "joinUrl": f"https://teams.microsoft.com/l/meetup-join/{thread_id}/0"
                        },
                    }
                ],
            }
        )

        _, result = await compose_resolve(client=client, permissions=full_permissions, url=url)
        assert "Standup" in result

    @pytest.mark.asyncio
    async def test_meeting_start_is_offset_aware(self, full_permissions):
        """UTC calendar event resolved via meetup-join URL: start/end must be
        offset-aware ISO with _timezone siblings."""
        url, thread_id = self._meeting_url()

        client = AsyncMock()
        client.get = AsyncMock(
            return_value={
                "value": [
                    {
                        "subject": "UTC Standup",
                        "start": {"dateTime": "2026-07-29T14:00:00.0000000", "timeZone": "UTC"},
                        "end": {"dateTime": "2026-07-29T14:30:00.0000000", "timeZone": "UTC"},
                        "organizer": {"emailAddress": {"name": "Avi"}},
                        "attendees": [],
                        "body": {"content": ""},
                        "location": {"displayName": ""},
                        "isOnlineMeeting": True,
                        "onlineMeeting": {
                            "joinUrl": f"https://teams.microsoft.com/l/meetup-join/{thread_id}/0"
                        },
                    }
                ],
            }
        )

        structured, _ = await compose_resolve(client=client, permissions=full_permissions, url=url)
        data = structured["data"]
        assert data["start"] == "2026-07-29T14:00:00+00:00"
        assert data["start_timezone"] == "UTC"
        assert data["end_timezone"] == "UTC"


# ---------------------------------------------------------------------------
# onedrive_file
# ---------------------------------------------------------------------------


class TestOneDriveFileResolve:
    @pytest.mark.asyncio
    async def test_onedrive_file_resolve(self, full_permissions):
        url = "https://contoso-my.sharepoint.com/personal/user_example_com/Documents/report.xlsx"

        client = AsyncMock()
        client.get = AsyncMock(
            return_value={
                "name": "report.xlsx",
                "size": 1024,
                "webUrl": "https://contoso-my.sharepoint.com/personal/user_example_com/Documents/report.xlsx",
            }
        )

        _, result = await compose_resolve(client=client, permissions=full_permissions, url=url)
        assert "report.xlsx" in result


# ---------------------------------------------------------------------------
# onedrive_share_link
# ---------------------------------------------------------------------------


class TestOneDriveShareLinkResolve:
    @pytest.mark.asyncio
    async def test_onedrive_share_link_resolve(self, full_permissions):
        url = "https://contoso-my.sharepoint.com/:x:/r/personal/user_example_com/_layouts/15/Doc.aspx?sourcedoc=%7Babc%7D"

        client = AsyncMock()
        client.get = AsyncMock(
            return_value={
                "name": "Roadmap.xlsx",
                "size": 2048,
                "webUrl": "https://contoso-my.sharepoint.com/personal/user_example_com/Shared%20Documents/Roadmap.xlsx",
            }
        )

        _, result = await compose_resolve(client=client, permissions=full_permissions, url=url)
        assert "Roadmap.xlsx" in result


# ---------------------------------------------------------------------------
# sharepoint_page
# ---------------------------------------------------------------------------


class TestSharePointPageResolve:
    @pytest.mark.asyncio
    async def test_sharepoint_page_found(self, full_permissions):
        url = "https://contoso.sharepoint.com/sites/MyProject/SitePages/Overview.aspx"

        site_response = {
            "id": "site-id-123",
            "displayName": "My Project",
            "webUrl": "https://contoso.sharepoint.com/sites/MyProject",
        }
        lists_response = {"value": [{"id": "list-id-456"}]}
        items_response = {
            "value": [
                {
                    "id": "42",
                    "webUrl": "https://contoso.sharepoint.com/sites/MyProject/SitePages/Overview.aspx",
                    "fields": {
                        "FileLeafRef": "Overview.aspx",
                        "Title": "Project Overview",
                        "Modified": "2026-01-01T00:00:00Z",
                    },
                }
            ]
        }

        client = AsyncMock()
        client.get = AsyncMock(side_effect=[site_response, lists_response, items_response])

        _, result = await compose_resolve(client=client, permissions=full_permissions, url=url)
        assert "SharePoint Page" in result
        assert "Project Overview" in result

    @pytest.mark.asyncio
    async def test_sharepoint_page_fallback(self, full_permissions):
        url = "https://contoso.sharepoint.com/sites/MyProject/SitePages/Overview.aspx"

        site_response = {
            "id": "site-id-123",
            "displayName": "My Project",
            "webUrl": "https://contoso.sharepoint.com/sites/MyProject",
        }

        client = AsyncMock()
        client.get = AsyncMock(
            side_effect=[
                site_response,
                GraphAPIError(404, "NotFound", "not found"),
            ]
        )

        _, result = await compose_resolve(client=client, permissions=full_permissions, url=url)
        assert "SharePoint Site" in result


# ---------------------------------------------------------------------------
# _get_event_by_id helper
# ---------------------------------------------------------------------------


class TestGetEventByIdHelper:
    @pytest.mark.asyncio
    async def test_returns_event_on_success(self):
        from ms365_intent_mcp.composers.resolve import _get_event_by_id

        client = AsyncMock()
        client.get = AsyncMock(
            return_value={
                "subject": "Project sync",
                "start": {"dateTime": "2026-05-26T10:00:00"},
                "end": {"dateTime": "2026-05-26T10:30:00"},
            }
        )
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
                {
                    "from": {"user": {"displayName": "Alice"}},
                    "body": {"content": "Hello"},
                    "createdDateTime": "2026-05-26T10:00:00Z",
                },
                {
                    "from": {"user": {"displayName": "Bob"}},
                    "body": {"content": "Hi"},
                    "createdDateTime": "2026-05-26T10:05:00Z",
                },
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
            _, result = await compose_resolve(
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
                return {
                    "value": [
                        {
                            **event_payload,
                            "onlineMeeting": {
                                "joinUrl": "https://teams.microsoft.com/l/meetup-join/19:meet@thread.v2/0"
                            },
                        }
                    ]
                }
            raise AssertionError(f"Unexpected endpoint: {endpoint}")

        client.get = AsyncMock(side_effect=fake_get)

        with patch("ms365_intent_mcp.composers.resolve.resolve_url") as mock_resolve:
            mock_resolve.return_value = ResolvedUrl(
                url_type="chat_thread",
                graph_endpoint="/chats/19:abc@thread.v2",
                required_scope="Chat.ReadWrite",
                extra={"chat_id": "19:abc@thread.v2"},
            )
            _, result = await compose_resolve(
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
            _, result = await compose_resolve(
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
            _, result = await compose_resolve(
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
            _, result = await compose_resolve(
                client=client,
                permissions=full_permissions,
                url="https://teams.microsoft.com/l/chat/19:abc@thread.v2/conversations",
            )

        assert "Project Sync" in result
        assert "⚠️" in result or "unavailable" in result.lower()
        assert "Hello" not in result  # messages failed; no message content should render

    @pytest.mark.asyncio
    async def test_chat_failure_partial_success(self, full_permissions, messages_payload):
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
            _, result = await compose_resolve(
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
            _, result = await compose_resolve(
                client=client,
                permissions=full_permissions,
                url="https://teams.microsoft.com/l/chat/19:abc@thread.v2/conversations",
            )

        # Chat + messages still render; meeting block silently absent
        assert "Project Sync" in result  # chat topic
        assert "Hello" in result

    @pytest.mark.asyncio
    async def test_messages_sorted_client_side(self, full_permissions, chat_meta, event_payload):
        out_of_order = {
            "value": [
                {
                    "from": {"user": {"displayName": "Alice"}},
                    "body": {"content": "OLDEST"},
                    "createdDateTime": "2026-05-26T08:00:00Z",
                },
                {
                    "from": {"user": {"displayName": "Bob"}},
                    "body": {"content": "NEWEST"},
                    "createdDateTime": "2026-05-26T10:00:00Z",
                },
                {
                    "from": {"user": {"displayName": "Carol"}},
                    "body": {"content": "MIDDLE"},
                    "createdDateTime": "2026-05-26T09:00:00Z",
                },
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
            _, result = await compose_resolve(
                client=client,
                permissions=full_permissions,
                url="https://teams.microsoft.com/l/chat/19:abc@thread.v2/conversations",
            )

        # Newest message must appear before middle and oldest
        newest_pos = result.find("NEWEST")
        middle_pos = result.find("MIDDLE")
        oldest_pos = result.find("OLDEST")
        assert 0 <= newest_pos < middle_pos < oldest_pos

    @pytest.mark.asyncio
    async def test_shared_recording_link_surfaced_end_to_end(
        self, full_permissions, chat_meta, event_payload
    ):
        """Issue #36: a recording shared in a chat must reach the rendered
        markdown as a fetchable link, not be dropped as unrecoverable media."""
        rec_url = (
            "https://tenant-my.sharepoint.com/personal/u/Documents/"
            "Microsoft Teams Chat Files/Bi-weekly-20260715_102544-Meeting Recording.mp4"
        )
        messages = {
            "value": [
                {
                    "from": {"user": {"displayName": "Bob"}},
                    "body": {
                        "content": (
                            "Here is a recording of what he has asked today,"
                            ' 15.07 10:25.<attachment id="a1"></attachment>'
                        )
                    },
                    "createdDateTime": "2026-07-15T10:25:00Z",
                    "attachments": [
                        {
                            "id": "a1",
                            "contentType": "reference",
                            "contentUrl": rec_url,
                            "content": None,
                            "name": "Bi-weekly-20260715_102544-Meeting Recording.mp4",
                        }
                    ],
                },
            ]
        }
        client = AsyncMock()

        async def fake_get(endpoint, params=None, headers=None):
            if endpoint == "/chats/19:abc@thread.v2":
                return chat_meta
            if endpoint.startswith("/chats/19:abc@thread.v2/messages"):
                return messages
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
            _, result = await compose_resolve(
                client=client,
                permissions=full_permissions,
                url="https://teams.microsoft.com/l/chat/19:abc@thread.v2/conversations",
            )

        assert "📎" in result
        assert rec_url in result
        assert "Meeting Recording.mp4" in result


class TestTruncateBody:
    """_truncate_body caps length but must never cut a markdown link mid-token,
    which would leave a dangling '[label](htt…'."""

    def test_short_text_unchanged(self):
        from ms365_intent_mcp.composers.resolve import _truncate_body

        assert _truncate_body("hello", 500) == "hello"

    def test_plain_text_truncated_with_ellipsis(self):
        from ms365_intent_mcp.composers.resolve import _truncate_body

        out = _truncate_body("x" * 600, 500)
        assert out == "x" * 500 + "…"

    def test_at_limit_no_ellipsis(self):
        from ms365_intent_mcp.composers.resolve import _truncate_body

        assert _truncate_body("x" * 500, 500) == "x" * 500

    def test_does_not_cut_link_mid_token(self):
        from ms365_intent_mcp.composers.resolve import _truncate_body

        # A link straddles the limit: cutting at 500 would land inside (url).
        pre = "y" * 490
        link = "[the doc](https://example.com/a/very/long/path/that/keeps/going)"
        out = _truncate_body(pre + link + " trailing", 500)
        # The link must be intact if present, or dropped entirely — never half.
        assert "](htt" not in out or out.count("[") == out.count("]")
        # Specifically: no dangling open-paren link.
        import re as _re

        assert not _re.search(r"\]\([^)]*$", out)

    def test_keeps_whole_link_when_it_fits_before_limit(self):
        from ms365_intent_mcp.composers.resolve import _truncate_body

        link = "[doc](https://x.com/y)"
        text = link + " " + "z" * 600
        out = _truncate_body(text, 500)
        assert link in out


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


class TestForwardedMessageExtraction:
    """Bug 3: messages with body=<attachment id=...> hide the real content
    in attachments[0].content.originalMessageContent (JSON string)."""

    def test_forwarded_message_extracts_inner_text(self):
        from ms365_intent_mcp.composers.resolve import _message_entry

        msg = {
            "createdDateTime": "2025-06-05T09:52:38Z",
            "from": {"user": {"id": "u1", "displayName": "Alice"}},
            "body": {"content": '<attachment id="123"></attachment>'},
            "attachments": [
                {
                    "id": "123",
                    "contentType": "forwardedMessageReference",
                    "content": '{"originalMessageContent": "<p>Hello <b>world</b></p>"}',
                }
            ],
        }
        entry = _message_entry(msg, {})
        assert entry["body"] == "Hello world"
        assert entry["is_body_empty"] is False

    def test_forwarded_message_truncates_long_inner_text(self):
        from ms365_intent_mcp.composers.resolve import _message_entry

        long_text = "x" * 600
        msg = {
            "createdDateTime": "2025-06-05T09:52:38Z",
            "from": {"user": {"id": "u1", "displayName": "Alice"}},
            "body": {"content": '<attachment id="123"></attachment>'},
            "attachments": [
                {
                    "contentType": "forwardedMessageReference",
                    "content": f'{{"originalMessageContent": "{long_text}"}}',
                }
            ],
        }
        entry = _message_entry(msg, {})
        assert entry["body"].endswith("…")
        assert len(entry["body"]) == 501

    def test_attachment_with_unparseable_content_falls_through(self):
        """Malformed JSON in attachment content must not crash; entry stays empty."""
        from ms365_intent_mcp.composers.resolve import _message_entry

        msg = {
            "createdDateTime": "2025-06-05T09:52:38Z",
            "from": {"user": {"id": "u1", "displayName": "Alice"}},
            "body": {"content": '<attachment id="123"></attachment>'},
            "attachments": [{"contentType": "forwardedMessageReference", "content": "{not json"}],
        }
        entry = _message_entry(msg, {})
        assert entry["is_body_empty"] is True

    def test_forwarded_non_dict_json_falls_through(self):
        """content that is valid JSON but not an object (list/string/number)
        must not crash — the entry degrades to empty body."""
        from ms365_intent_mcp.composers.resolve import _message_entry

        for payload in ("[1,2,3]", '"just a string"', "42"):
            msg = {
                "createdDateTime": "2025-06-05T09:52:38Z",
                "from": {"user": {"id": "u1", "displayName": "Alice"}},
                "body": {"content": '<attachment id="123"></attachment>'},
                "attachments": [{"contentType": "forwardedMessageReference", "content": payload}],
            }
            entry = _message_entry(msg, {})
            assert entry["is_body_empty"] is True

    def test_non_forwarded_attachment_does_not_inject_body(self):
        """An image/file attachment without 'forwardedMessageReference' should
        not be misread as a forwarded message — body stays empty."""
        from ms365_intent_mcp.composers.resolve import _message_entry

        msg = {
            "createdDateTime": "2025-06-05T09:52:38Z",
            "from": {"user": {"id": "u1", "displayName": "Alice"}},
            "body": {"content": '<attachment id="123"></attachment>'},
            "attachments": [{"contentType": "image/png", "name": "foo.png"}],
        }
        entry = _message_entry(msg, {})
        assert entry["is_body_empty"] is True

    def test_real_body_takes_precedence_over_attachment(self):
        """If body has stripped text, use that — don't override with attachment content."""
        from ms365_intent_mcp.composers.resolve import _message_entry

        msg = {
            "createdDateTime": "2025-06-05T09:52:38Z",
            "from": {"user": {"id": "u1", "displayName": "Alice"}},
            "body": {"content": "<p>Real body</p>"},
            "attachments": [
                {
                    "contentType": "forwardedMessageReference",
                    "content": '{"originalMessageContent": "<p>Forwarded</p>"}',
                }
            ],
        }
        entry = _message_entry(msg, {})
        assert entry["body"] == "Real body"


class TestReferenceAttachmentExtraction:
    """Issue #36: shared files/recordings arrive as reference-type attachments
    (contentType='reference', name + contentUrl). These must be surfaced, not
    silently dropped — otherwise a fetchable recording reads as unrecoverable
    pasted media."""

    def test_reference_attachment_surfaced(self):
        from ms365_intent_mcp.composers.resolve import _message_entry

        msg = {
            "createdDateTime": "2026-07-15T10:25:00Z",
            "from": {"user": {"id": "u1", "displayName": "Alice"}},
            "body": {"content": '<div><attachment id="abc"></attachment></div>'},
            "attachments": [
                {
                    "id": "abc",
                    "contentType": "reference",
                    "contentUrl": "https://tenant-my.sharepoint.com/personal/u/Documents/Microsoft%20Teams%20Chat%20Files/Meeting%20Recording.mp4",
                    "content": None,
                    "name": "Meeting Recording.mp4",
                }
            ],
        }
        entry = _message_entry(msg, {})
        assert entry["attachments"] == [
            {
                "name": "Meeting Recording.mp4",
                "url": "https://tenant-my.sharepoint.com/personal/u/Documents/Microsoft%20Teams%20Chat%20Files/Meeting%20Recording.mp4",
            }
        ]

    def test_message_with_only_attachment_still_carries_it(self):
        """Body is just the <attachment> placeholder (empty after strip); the
        file must still surface even though is_body_empty is True."""
        from ms365_intent_mcp.composers.resolve import _message_entry

        msg = {
            "createdDateTime": "2026-07-15T10:25:00Z",
            "from": {"user": {"id": "u1", "displayName": "Alice"}},
            "body": {"content": '<div><attachment id="abc"></attachment></div>'},
            "attachments": [
                {
                    "id": "abc",
                    "contentType": "reference",
                    "contentUrl": "https://tenant-my.sharepoint.com/personal/u/rec.mp4",
                    "name": "rec.mp4",
                }
            ],
        }
        entry = _message_entry(msg, {})
        assert entry["is_body_empty"] is True
        assert len(entry["attachments"]) == 1
        assert entry["attachments"][0]["url"].endswith("rec.mp4")

    def test_message_with_text_and_attachment(self):
        from ms365_intent_mcp.composers.resolve import _message_entry

        msg = {
            "createdDateTime": "2026-07-15T10:25:00Z",
            "from": {"user": {"id": "u1", "displayName": "Alice"}},
            "body": {"content": '<p>Here is the recording</p><attachment id="abc"></attachment>'},
            "attachments": [
                {
                    "contentType": "reference",
                    "contentUrl": "https://tenant-my.sharepoint.com/personal/u/rec.mp4",
                    "name": "rec.mp4",
                }
            ],
        }
        entry = _message_entry(msg, {})
        assert entry["body"] == "Here is the recording"
        assert entry["is_body_empty"] is False
        assert entry["attachments"] == [
            {"name": "rec.mp4", "url": "https://tenant-my.sharepoint.com/personal/u/rec.mp4"}
        ]

    def test_multiple_reference_attachments(self):
        from ms365_intent_mcp.composers.resolve import _message_entry

        msg = {
            "createdDateTime": "2026-07-15T10:25:00Z",
            "from": {"user": {"id": "u1", "displayName": "Alice"}},
            "body": {"content": "<div></div>"},
            "attachments": [
                {"contentType": "reference", "contentUrl": "https://x/a.mp4", "name": "a.mp4"},
                {"contentType": "reference", "contentUrl": "https://x/b.pdf", "name": "b.pdf"},
            ],
        }
        entry = _message_entry(msg, {})
        assert len(entry["attachments"]) == 2
        assert entry["attachments"][0]["name"] == "a.mp4"
        assert entry["attachments"][1]["name"] == "b.pdf"

    def test_no_attachments_yields_empty_list(self):
        from ms365_intent_mcp.composers.resolve import _message_entry

        msg = {
            "createdDateTime": "2026-07-15T10:25:00Z",
            "from": {"user": {"id": "u1", "displayName": "Alice"}},
            "body": {"content": "<p>hi</p>"},
        }
        entry = _message_entry(msg, {})
        assert entry["attachments"] == []

    def test_forwarded_reference_only_not_treated_as_file(self):
        """forwardedMessageReference is not a shared file — it must not leak
        into the attachments list."""
        from ms365_intent_mcp.composers.resolve import _message_entry

        msg = {
            "createdDateTime": "2026-07-15T10:25:00Z",
            "from": {"user": {"id": "u1", "displayName": "Alice"}},
            "body": {"content": '<attachment id="123"></attachment>'},
            "attachments": [
                {
                    "contentType": "forwardedMessageReference",
                    "content": '{"originalMessageContent": "<p>fwd</p>"}',
                }
            ],
        }
        entry = _message_entry(msg, {})
        assert entry["attachments"] == []

    def test_reference_without_url_skipped(self):
        """A reference attachment missing contentUrl is not renderable — skip it."""
        from ms365_intent_mcp.composers.resolve import _message_entry

        msg = {
            "createdDateTime": "2026-07-15T10:25:00Z",
            "from": {"user": {"id": "u1", "displayName": "Alice"}},
            "body": {"content": "<div></div>"},
            "attachments": [{"contentType": "reference", "name": "orphan.mp4"}],
        }
        entry = _message_entry(msg, {})
        assert entry["attachments"] == []


class TestReplyContextExtraction:
    """A reply/quote surfaces as a messageReference attachment whose `content`
    is a JSON string with messagePreview + messageSender. Surface it as context
    so the reply is readable on its own (user request on issue #36)."""

    def _reply_msg(self, sender="Bob", preview="the original question"):
        import json as _json

        return {
            "createdDateTime": "2026-07-15T10:25:00Z",
            "from": {"user": {"id": "u1", "displayName": "Alice"}},
            "body": {"content": "<p>Yes, agreed</p>"},
            "attachments": [
                {
                    "id": "1728422677844",
                    "contentType": "messageReference",
                    "contentUrl": None,
                    "content": _json.dumps(
                        {
                            "messageId": "1728422677844",
                            "messagePreview": preview,
                            "messageSender": {"user": {"id": "x", "displayName": sender}},
                        }
                    ),
                }
            ],
        }

    def test_reply_context_extracted(self):
        from ms365_intent_mcp.composers.resolve import _message_entry

        entry = _message_entry(self._reply_msg(), {})
        assert entry["body"] == "Yes, agreed"
        assert entry["reply_to"] == {
            "sender": "Bob",
            "preview": "the original question",
        }

    def test_reply_context_absent_yields_none(self):
        from ms365_intent_mcp.composers.resolve import _message_entry

        msg = {
            "createdDateTime": "2026-07-15T10:25:00Z",
            "from": {"user": {"id": "u1", "displayName": "Alice"}},
            "body": {"content": "<p>hi</p>"},
        }
        entry = _message_entry(msg, {})
        assert entry["reply_to"] is None

    def test_reply_context_unparseable_content_yields_none(self):
        from ms365_intent_mcp.composers.resolve import _message_entry

        msg = {
            "createdDateTime": "2026-07-15T10:25:00Z",
            "from": {"user": {"id": "u1", "displayName": "Alice"}},
            "body": {"content": "<p>hi</p>"},
            "attachments": [{"contentType": "messageReference", "content": "{not json"}],
        }
        entry = _message_entry(msg, {})
        assert entry["reply_to"] is None

    def test_reply_context_non_dict_json_yields_none(self):
        """content that is valid JSON but not an object (list/string/number)
        must not crash — a malformed reference should degrade to None."""
        from ms365_intent_mcp.composers.resolve import _message_entry

        for payload in ("[1,2,3]", '"just a string"', "42"):
            msg = {
                "createdDateTime": "2026-07-15T10:25:00Z",
                "from": {"user": {"id": "u1", "displayName": "Alice"}},
                "body": {"content": "<p>hi</p>"},
                "attachments": [{"contentType": "messageReference", "content": payload}],
            }
            entry = _message_entry(msg, {})
            assert entry["reply_to"] is None

    def test_reply_preview_truncated(self):
        from ms365_intent_mcp.composers.resolve import _message_entry

        entry = _message_entry(self._reply_msg(preview="x" * 300), {})
        assert entry["reply_to"]["preview"].endswith("…")
        assert len(entry["reply_to"]["preview"]) == 201

    def test_reply_sender_falls_back_when_missing(self):
        import json as _json

        from ms365_intent_mcp.composers.resolve import _message_entry

        msg = {
            "createdDateTime": "2026-07-15T10:25:00Z",
            "from": {"user": {"id": "u1", "displayName": "Alice"}},
            "body": {"content": "<p>reply</p>"},
            "attachments": [
                {
                    "contentType": "messageReference",
                    "content": _json.dumps({"messagePreview": "context"}),
                }
            ],
        }
        entry = _message_entry(msg, {})
        assert entry["reply_to"] == {"sender": "Unknown", "preview": "context"}


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
            self._recording_event(
                "c1",
                "2026-05-29T10:25:00Z",
                "success",
                url="https://recording.example/c1.mp4",
                duration="PT25M0S",
            ),
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
            self._recording_event(
                "c1", "2026-05-29T10:00:00Z", "success", url="https://r1", duration="PT5M0S"
            ),
            self._recording_event(
                "c2", "2026-05-29T11:00:00Z", "success", url="https://r2", duration="PT10M0S"
            ),
        ]
        out = _group_call_events(events, {})
        assert len(out) == 2
        urls = {c["recording_url"] for c in out}
        assert urls == {"https://r1", "https://r2"}

    def test_recording_url_prefers_success_status(self):
        from ms365_intent_mcp.composers.resolve import _group_call_events

        events = [
            self._recording_event(
                "c1", "2026-05-29T10:00:00Z", "chunkFinished", url="https://chunk.example/temp.mp4"
            ),
            self._recording_event(
                "c1", "2026-05-29T10:25:00Z", "success", url="https://final.example/c1.mp4"
            ),
        ]
        out = _group_call_events(events, {})
        assert out[0]["recording_url"] == "https://final.example/c1.mp4"

    def test_recording_url_pending_no_success(self):
        from ms365_intent_mcp.composers.resolve import _group_call_events

        events = [
            self._recording_event("c1", "2026-05-29T10:00:00Z", "initial"),
            self._recording_event(
                "c1", "2026-05-29T10:05:00Z", "chunkFinished", url="https://chunk.example/temp.mp4"
            ),
        ]
        out = _group_call_events(events, {})
        assert out[0]["recording_url"] == ""

    def test_status_case_insensitive(self):
        from ms365_intent_mcp.composers.resolve import _group_call_events

        events = [
            self._recording_event(
                "c1", "2026-05-29T10:00:00Z", "Success", url="https://final.example/c1.mp4"
            ),
        ]
        out = _group_call_events(events, {})
        assert out[0]["recording_url"] == "https://final.example/c1.mp4"

    def test_transcript_ready_flag(self):
        from ms365_intent_mcp.composers.resolve import _group_call_events

        events = [
            self._recording_event("c1", "2026-05-29T10:00:00Z", "success", url="https://r1"),
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

        events = [self._recording_event("c1", "2026-05-29T10:00:00Z", "success", url="https://r1")]
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
            self._recording_event(
                "c1", "2026-05-27T12:12:05Z", "chunkFinished", duration="PT25M38S"
            ),
            self._recording_event(
                "c1", "2026-05-27T12:12:26Z", "success", url="https://r1", duration="PT25M38S"
            ),
        ]
        out = _group_call_events(events, {})
        assert out[0]["duration"] == "25m38s"

    def test_duration_falls_back_to_chunk_when_no_success(self):
        """Recording in progress: success event hasn't fired yet. Use the
        latest non-zero chunk duration instead of the initial PT0S."""
        from ms365_intent_mcp.composers.resolve import _group_call_events

        events = [
            self._recording_event("c1", "2026-05-27T11:46:27Z", "initial", duration="PT0S"),
            self._recording_event(
                "c1", "2026-05-27T12:12:05Z", "chunkFinished", duration="PT25M38S"
            ),
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


class TestEncodeShareUrl:
    def test_produces_u_prefix_base64url(self):
        from ms365_intent_mcp.composers.resolve import _encode_share_url

        url = "https://sap-my.sharepoint.com/:v:/p/marcus_karlbowski/IQBw"
        encoded = _encode_share_url(url)
        assert encoded.startswith("u!")
        # Base64url has no padding.
        assert "=" not in encoded


class TestExtractRecordingOwner:
    def test_extracts_upn_from_share_url(self):
        from ms365_intent_mcp.composers.resolve import _extract_recording_owner

        url = "https://sap-my.sharepoint.com/:v:/p/marcus_karlbowski/IQBw"
        host, upn = _extract_recording_owner(url)
        assert host == "sap-my.sharepoint.com"
        assert upn == "marcus_karlbowski"

    def test_returns_empty_when_shape_unrecognized(self):
        from ms365_intent_mcp.composers.resolve import _extract_recording_owner

        host, upn = _extract_recording_owner("https://example.com/nope")
        assert host == ""
        assert upn == ""

    def test_returns_empty_on_empty_input(self):
        from ms365_intent_mcp.composers.resolve import _extract_recording_owner

        assert _extract_recording_owner("") == ("", "")


class TestEnrichCallRecording:
    @pytest.mark.asyncio
    async def test_adds_drive_metadata_when_shares_succeeds(self):
        from ms365_intent_mcp.composers.resolve import _enrich_call_recording

        client = AsyncMock()
        client.get = AsyncMock(
            return_value={
                "id": "01LUWJL4TQYSFUZ4AUBVC334ND5VUPJEVP",
                "parentReference": {"driveId": "b!vwRb..."},
            }
        )

        entry = {
            "kind": "call",
            "recording_url": "https://sap-my.sharepoint.com/:v:/p/marcus_karlbowski/IQBw",
        }
        await _enrich_call_recording(client, entry)

        assert entry["drive_id"] == "b!vwRb..."
        assert entry["drive_item_id"] == "01LUWJL4TQYSFUZ4AUBVC334ND5VUPJEVP"
        assert entry["owner_upn"] == "marcus_karlbowski"
        assert entry["vroom_url"] == (
            "https://sap-my.sharepoint.com/personal/marcus_karlbowski"
            "/_api/v2.0/drives/b!vwRb.../items/01LUWJL4TQYSFUZ4AUBVC334ND5VUPJEVP"
        )
        client.get.assert_called_once()
        called_endpoint = client.get.call_args[0][0]
        assert called_endpoint.startswith("/shares/u!")
        assert called_endpoint.endswith("/driveItem")

    @pytest.mark.asyncio
    async def test_no_op_when_recording_url_missing(self):
        from ms365_intent_mcp.composers.resolve import _enrich_call_recording

        client = AsyncMock()
        entry = {"kind": "call", "recording_url": ""}
        await _enrich_call_recording(client, entry)
        client.get.assert_not_called()
        assert "drive_id" not in entry
        assert "vroom_url" not in entry

    @pytest.mark.asyncio
    async def test_silent_degradation_on_shares_403(self):
        """Cross-organizer recordings may 403 until the user opens the
        recording in Teams UI. Enrichment must not raise or overwrite the URL —
        it just leaves the drive fields absent."""
        from ms365_intent_mcp.composers.resolve import _enrich_call_recording

        client = AsyncMock()
        client.get = AsyncMock(side_effect=GraphAPIError(403, "Forbidden", "no access"))

        entry = {
            "kind": "call",
            "recording_url": "https://sap-my.sharepoint.com/:v:/p/other_user/XYZ",
        }
        await _enrich_call_recording(client, entry)

        assert entry["recording_url"] == "https://sap-my.sharepoint.com/:v:/p/other_user/XYZ"
        assert "drive_id" not in entry
        assert "vroom_url" not in entry

    @pytest.mark.asyncio
    async def test_no_vroom_when_upn_not_extractable(self):
        """If URL shape doesn't yield an owner UPN, drive_id/item_id are still
        exposed — vroom_url just isn't composable."""
        from ms365_intent_mcp.composers.resolve import _enrich_call_recording

        client = AsyncMock()
        client.get = AsyncMock(
            return_value={
                "id": "01ABC",
                "parentReference": {"driveId": "b!x"},
            }
        )
        entry = {"kind": "call", "recording_url": "https://example.com/no-upn-here"}
        await _enrich_call_recording(client, entry)

        assert entry.get("drive_id") == "b!x"
        assert entry.get("drive_item_id") == "01ABC"
        assert entry.get("owner_upn") == ""
        assert "vroom_url" not in entry


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
        client.get = AsyncMock(
            return_value={
                "value": [{"id": "m1"}, {"id": "m2"}],
            }
        )
        messages, err = await _paginate_chat_messages(client, "chat1")
        assert len(messages) == 2
        assert err is None

    @pytest.mark.asyncio
    async def test_follows_next_link(self):
        from ms365_intent_mcp.composers.resolve import _paginate_chat_messages

        client = AsyncMock()
        responses = [
            {
                "value": [{"id": "m1"}, {"id": "m2"}],
                "@odata.nextLink": "https://graph.microsoft.com/v1.0/chats/chat1/messages?$skiptoken=p2",
            },
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
            {
                "value": [{"id": "m1"}, {"id": "m2"}],
                "@odata.nextLink": "https://graph.microsoft.com/v1.0/chats/chat1/messages?$skiptoken=p2",
            },
            {"value": [{"id": "m2"}, {"id": "m3"}]},  # m2 duplicates
        ]
        client.get = AsyncMock(side_effect=responses)
        messages, err = await _paginate_chat_messages(client, "chat1")
        assert [m["id"] for m in messages] == ["m1", "m2", "m3"]

    @pytest.mark.asyncio
    async def test_includes_messages_without_id(self):
        from ms365_intent_mcp.composers.resolve import _paginate_chat_messages

        client = AsyncMock()
        client.get = AsyncMock(
            return_value={
                "value": [{"id": "m1"}, {"createdDateTime": "x"}],  # second has no id
            }
        )
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
            {
                "value": [{"id": "m1"}, {"id": "m2"}],
                "@odata.nextLink": "https://graph.microsoft.com/v1.0/chats/chat1/messages?$skiptoken=p2",
            },
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
        page1 = {
            "value": [{"id": f"m{i}"} for i in range(60)],
            "@odata.nextLink": "https://graph.microsoft.com/v1.0/chats/chat1/messages?$skiptoken=p2",
        }
        page2 = {"value": [{"id": f"m{60 + i}"} for i in range(60)]}
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


# ---------------------------------------------------------------------------
# Email attachment wiring
# ---------------------------------------------------------------------------


class TestResolveEmailAttachments:
    @pytest.mark.asyncio
    async def test_enumerates_when_has_attachments(self, full_permissions):
        client = AsyncMock()
        client.get = AsyncMock(
            side_effect=[
                {  # message fetch
                    "subject": "Bug report",
                    "from": {"emailAddress": {"name": "Cust"}},
                    "receivedDateTime": "2026-07-20T08:00:00Z",
                    "body": {"contentType": "text", "content": "see [cid:img1@01DD]"},
                    "hasAttachments": True,
                },
                {  # /attachments
                    "value": [
                        {
                            "@odata.type": "#microsoft.graph.fileAttachment",
                            "name": "shot.png",
                            "contentType": "image/png",
                            "size": 12,
                            "isInline": True,
                            "contentId": "img1@01DD",
                            "contentBytes": "AA==",
                            "id": "AT1",
                        }
                    ],
                },
            ]
        )
        with patch("ms365_intent_mcp.composers.resolve.resolve_url") as mock_resolve:
            mock_resolve.return_value = ResolvedUrl(
                url_type="email",
                graph_endpoint="/me/messages/M1",
                required_scope="Mail.Read",
            )
            data, md = await compose_resolve(
                client=client,
                permissions=full_permissions,
                url="https://outlook.office365.com/mail/id/M1",
            )
        atts = data["data"]["attachments"]
        assert len(atts) == 1
        assert atts[0]["cid"] == "img1@01DD"
        assert atts[0]["local_path"] is None  # no output_dir → metadata only
        assert "_content_bytes" not in atts[0]  # internal field stripped
        assert "shot.png" in md

    @pytest.mark.asyncio
    async def test_no_extra_call_when_no_attachments_no_cid(self, full_permissions):
        client = AsyncMock()
        client.get = AsyncMock(
            return_value={
                "subject": "plain",
                "from": {"emailAddress": {"name": "A"}},
                "receivedDateTime": "2026-07-20T08:00:00Z",
                "body": {"contentType": "text", "content": "no images here"},
                "hasAttachments": False,
            }
        )
        with patch("ms365_intent_mcp.composers.resolve.resolve_url") as mock_resolve:
            mock_resolve.return_value = ResolvedUrl(
                url_type="email",
                graph_endpoint="/me/messages/M2",
                required_scope="Mail.Read",
            )
            data, _ = await compose_resolve(
                client=client,
                permissions=full_permissions,
                url="https://outlook.office365.com/mail/id/M2",
            )
        assert client.get.await_count == 1  # message only, no /attachments
        assert data["data"]["attachments"] == []

    @pytest.mark.asyncio
    async def test_downloads_when_output_dir_given(self, full_permissions, tmp_path):
        import base64

        client = AsyncMock()
        client.get = AsyncMock(
            side_effect=[
                {
                    "subject": "s",
                    "from": {"emailAddress": {"name": "A"}},
                    "receivedDateTime": "2026-07-20T08:00:00Z",
                    "body": {"contentType": "text", "content": "x"},
                    "hasAttachments": True,
                },
                {
                    "value": [
                        {
                            "@odata.type": "#microsoft.graph.fileAttachment",
                            "name": "r.pdf",
                            "contentType": "application/pdf",
                            "size": 5,
                            "isInline": False,
                            "contentBytes": base64.b64encode(b"hello").decode(),
                            "id": "AT2",
                        }
                    ]
                },
            ]
        )
        with patch("ms365_intent_mcp.composers.resolve.resolve_url") as mock_resolve:
            mock_resolve.return_value = ResolvedUrl(
                url_type="email",
                graph_endpoint="/me/messages/M3",
                required_scope="Mail.Read",
            )
            data, _ = await compose_resolve(
                client=client,
                permissions=full_permissions,
                url="https://outlook.office365.com/mail/id/M3",
                output_dir=str(tmp_path),
            )
        lp = data["data"]["attachments"][0]["local_path"]
        assert lp is not None
        from pathlib import Path

        assert Path(lp).read_bytes() == b"hello"

    @pytest.mark.asyncio
    async def test_enumeration_error_degrades(self, full_permissions):
        client = AsyncMock()
        client.get = AsyncMock(
            side_effect=[
                {
                    "subject": "s",
                    "from": {"emailAddress": {"name": "A"}},
                    "receivedDateTime": "2026-07-20T08:00:00Z",
                    "body": {"contentType": "text", "content": "x"},
                    "hasAttachments": True,
                },
                GraphAPIError(403, "ErrorAccessDenied", "no"),
            ]
        )
        with patch("ms365_intent_mcp.composers.resolve.resolve_url") as mock_resolve:
            mock_resolve.return_value = ResolvedUrl(
                url_type="email",
                graph_endpoint="/me/messages/M4",
                required_scope="Mail.Read",
            )
            data, md = await compose_resolve(
                client=client,
                permissions=full_permissions,
                url="https://outlook.office365.com/mail/id/M4",
            )
        assert data["data"]["attachments"] == []
        assert "s" in md  # body still renders (subject present)

    @pytest.mark.asyncio
    async def test_non_graph_error_in_download_degrades_gracefully(
        self, full_permissions, tmp_path
    ):
        """Non-GraphAPIError (e.g. RuntimeError, httpx.TimeoutException) from
        download_attachments must not propagate — body + subject still render."""
        import base64

        client = AsyncMock()
        client.get = AsyncMock(
            side_effect=[
                {
                    "subject": "Crash report",
                    "from": {"emailAddress": {"name": "B"}},
                    "receivedDateTime": "2026-07-20T09:00:00Z",
                    "body": {"contentType": "text", "content": "body text here"},
                    "hasAttachments": True,
                },
                {
                    "value": [
                        {
                            "@odata.type": "#microsoft.graph.fileAttachment",
                            "name": "crash.log",
                            "contentType": "text/plain",
                            "size": 3,
                            "isInline": False,
                            "contentBytes": base64.b64encode(b"err").decode(),
                            "id": "AT9",
                        }
                    ]
                },
            ]
        )
        with (
            patch("ms365_intent_mcp.composers.resolve.resolve_url") as mock_resolve,
            patch(
                "ms365_intent_mcp.composers.resolve.download_attachments",
                side_effect=RuntimeError("boom"),
            ) as mock_dl,
        ):
            mock_resolve.return_value = ResolvedUrl(
                url_type="email",
                graph_endpoint="/me/messages/M5",
                required_scope="Mail.Read",
            )
            data, md = await compose_resolve(
                client=client,
                permissions=full_permissions,
                url="https://outlook.office365.com/mail/id/M5",
                output_dir=str(tmp_path),
            )
        mock_dl.assert_called_once()
        assert "Crash report" in md
        assert "body text here" in md or "B" in md
        assert data["data"]["subject"] == "Crash report"
