"""Tests for resolve composer."""

from unittest.mock import AsyncMock, patch

import pytest

from ms365_intent_mcp.composers.resolve import compose_resolve
from ms365_intent_mcp.graph import GraphAPIError
from ms365_intent_mcp.permissions import PermissionRegistry
from ms365_intent_mcp.resolver import UrlParseError


@pytest.fixture
def full_permissions():
    return PermissionRegistry(["Mail.Read", "Chat.Read", "Calendars.Read", "Files.Read", "Sites.Read.All"])


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
            from ms365_intent_mcp.resolver import ResolvedUrl
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
            from ms365_intent_mcp.resolver import ResolvedUrl
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
            from ms365_intent_mcp.resolver import ResolvedUrl
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
