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
        "Chat.Read",
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
        page_response = {
            "name": "Overview.aspx",
            "webUrl": "https://sap.sharepoint.com/sites/MyProject/SitePages/Overview.aspx",
            "lastModifiedDateTime": "2026-01-01T00:00:00Z",
            "size": 512,
        }

        client = AsyncMock()
        client.get = AsyncMock(side_effect=[site_response, page_response])

        result = await compose_resolve(client=client, permissions=full_permissions, url=url)
        assert "SharePoint Page" in result
        assert "Overview.aspx" in result

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
