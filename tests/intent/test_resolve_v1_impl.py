"""Unit tests for _resolve_v1_impl — mocked context, no FastMCP."""

from unittest.mock import AsyncMock, MagicMock

import pytest

from ms365_intent_mcp.graph import GraphAPIError
from ms365_intent_mcp.intent._shared import ErrorResponse
from ms365_intent_mcp.intent.resolve.impl import _resolve_v1_impl
from ms365_intent_mcp.intent.resolve.schemas import (
    ChatThreadContent,
    EmailContent,
    OneDriveFileContent,
    ResolvePayload,
    ResolvedContent,
    SharePointPageContent,
)


def _mock_ctx():
    """Build a mocked FastMCP Context with the three deps the impls need."""
    ctx = MagicMock()
    config = MagicMock(default_timezone="Europe/Berlin")
    client = AsyncMock()
    permissions = MagicMock()
    permissions.check = MagicMock(return_value=None)
    ctx.request_context.lifespan_context = {
        "config": config,
        "client": client,
        "permissions": permissions,
    }
    return ctx, client, permissions


class TestResolveV1HappyPath:
    @pytest.mark.asyncio
    async def test_email_url_returns_email_content(self, monkeypatch):
        ctx, _, _ = _mock_ctx()

        async def _fake(client, permissions, url):
            return {"url": url, "kind": "email", "data": {"kind": "email", "subject": "Test Email", "sender": "sender@example.com", "body": ""}}, "**Subject:** Test Email\nFrom: sender@example.com"

        monkeypatch.setattr(
            "ms365_intent_mcp.intent.resolve.impl.compose_resolve",
            _fake,
        )

        payload = ResolvePayload.model_validate({
            "url": "https://outlook.office.com/mail/id/AAMkAGNhYWU5ZjBhLTQ4YjQtNGViNi1hZmM0LTJhYmJhNGE0YjFlNgBGAAAAAABmzX8NV4RzQqSJsepvY8W8BwARKi4ZYdHrQ76rWR9vLcK2AAAAAAEMAAARKi4ZYdHrQ76rWR9vLcK2AABQN3eTAAA%3D"
        })
        response = await _resolve_v1_impl(ctx, payload)

        assert isinstance(response, ResolvedContent)
        assert response.type == "resolved_content"
        assert response.kind == "email"
        assert isinstance(response.data, EmailContent)
        assert response.data.kind == "email"
        assert "Subject" in response.rendered_markdown

    @pytest.mark.asyncio
    async def test_chat_thread_url_returns_chat_thread_content(self, monkeypatch):
        ctx, _, _ = _mock_ctx()

        async def _fake(client, permissions, url):
            return {"url": url, "kind": "chat_thread", "data": {"kind": "chat_thread"}}, "**Chat Thread**\n5 recent messages"

        monkeypatch.setattr(
            "ms365_intent_mcp.intent.resolve.impl.compose_resolve",
            _fake,
        )

        payload = ResolvePayload.model_validate({
            "url": "https://teams.microsoft.com/l/chat/19:abc123def456@thread.v2/0"
        })
        response = await _resolve_v1_impl(ctx, payload)

        assert isinstance(response, ResolvedContent)
        assert response.kind == "chat_thread"
        assert isinstance(response.data, ChatThreadContent)
        assert response.data.kind == "chat_thread"

    @pytest.mark.asyncio
    async def test_sharepoint_page_returns_sharepoint_content(self, monkeypatch):
        ctx, _, _ = _mock_ctx()

        async def _fake(client, permissions, url):
            return {"url": url, "kind": "sharepoint_page", "data": {"kind": "sharepoint_page", "title": "Engineering Wiki"}}, "**Page:** Engineering Wiki\nLast modified: 2026-07-01"

        monkeypatch.setattr(
            "ms365_intent_mcp.intent.resolve.impl.compose_resolve",
            _fake,
        )

        payload = ResolvePayload.model_validate({
            "url": "https://contoso.sharepoint.com/sites/Engineering/SitePages/Home.aspx"
        })
        response = await _resolve_v1_impl(ctx, payload)

        assert isinstance(response, ResolvedContent)
        assert response.kind == "sharepoint_page"
        assert isinstance(response.data, SharePointPageContent)
        assert response.data.kind == "sharepoint_page"

    @pytest.mark.asyncio
    async def test_onedrive_share_link_returns_onedrive_content(self, monkeypatch):
        ctx, _, _ = _mock_ctx()

        async def _fake(client, permissions, url):
            return {"url": url, "kind": "onedrive_share_link", "data": {"kind": "onedrive_file", "name": "document.docx"}}, "**File:** document.docx\nSize: 1.2 MB"

        monkeypatch.setattr(
            "ms365_intent_mcp.intent.resolve.impl.compose_resolve",
            _fake,
        )

        payload = ResolvePayload.model_validate({
            "url": "https://contoso.sharepoint.com/personal/alice_contoso_com/_layouts/15/Doc.aspx?sourcedoc=%7Babc%7D"
        })
        response = await _resolve_v1_impl(ctx, payload)

        assert isinstance(response, ResolvedContent)
        assert response.kind == "onedrive_file"
        assert isinstance(response.data, OneDriveFileContent)
        assert response.data.kind == "onedrive_file"

    @pytest.mark.asyncio
    async def test_rendered_markdown_passthrough(self, monkeypatch):
        """The markdown from compose_resolve is surfaced verbatim in rendered_markdown."""
        ctx, _, _ = _mock_ctx()
        expected_markdown = "## My Resolved Content\nFull details here."

        async def _fake(client, permissions, url):
            return {"url": url, "kind": "email", "data": {"kind": "email", "subject": "x", "sender": "a", "body": ""}}, expected_markdown

        monkeypatch.setattr(
            "ms365_intent_mcp.intent.resolve.impl.compose_resolve",
            _fake,
        )

        payload = ResolvePayload.model_validate({
            "url": "https://outlook.office.com/mail/id/AAMkAGNhYWU5ZjBhLTQ4YjQtNGViNi1hZmM0LTJhYmJhNGE0YjFlNgBGAAAAAABmzX8NV4RzQqSJsepvY8W8BwARKi4ZYdHrQ76rWR9vLcK2AAAAAAEMAAARKi4ZYdHrQ76rWR9vLcK2AABQN3eTAAA%3D"
        })
        response = await _resolve_v1_impl(ctx, payload)

        assert response.rendered_markdown == expected_markdown


    @pytest.mark.asyncio
    async def test_invalid_structured_data_returns_validation_error(self, monkeypatch):
        """Composer returning partial/invalid structured_data must yield ErrorResponse(code='validation_error')."""
        ctx, _, _ = _mock_ctx()

        async def _fake(client, permissions, url):
            # 'email' kind requires 'subject', 'sender', 'body' — return none of them
            return {"url": url, "kind": "email", "data": {"totally": "wrong"}}, "partial"

        monkeypatch.setattr(
            "ms365_intent_mcp.intent.resolve.impl.compose_resolve",
            _fake,
        )

        payload = ResolvePayload.model_validate({
            "url": "https://outlook.office.com/mail/id/AAMkAGNhYWU5ZjBhLTQ4YjQtNGViNi1hZmM0LTJhYmJhNGE0YjFlNgBGAAAAAABmzX8NV4RzQqSJsepvY8W8BwARKi4ZYdHrQ76rWR9vLcK2AAAAAAEMAAARKi4ZYdHrQ76rWR9vLcK2AABQN3eTAAA%3D"
        })
        response = await _resolve_v1_impl(ctx, payload)

        assert isinstance(response, ErrorResponse)
        assert response.code == "validation_error"
        assert "email" in response.message
    @pytest.mark.asyncio
    async def test_invalid_url_returns_error_response(self, monkeypatch):
        """An unrecognised URL should yield ErrorResponse with code='invalid_id'."""
        ctx, _, _ = _mock_ctx()

        payload = ResolvePayload.model_validate({
            "url": "https://example.com/not-an-m365-url"
        })
        response = await _resolve_v1_impl(ctx, payload)

        assert isinstance(response, ErrorResponse)
        assert response.type == "error"
        assert response.code == "invalid_id"

    @pytest.mark.asyncio
    async def test_graph_api_error_returns_error_response(self, monkeypatch):
        """GraphAPIError from compose_resolve should yield ErrorResponse with code='graph_api_error'."""
        ctx, _, _ = _mock_ctx()

        async def _fake(client, permissions, url):
            raise GraphAPIError(status_code=503, error_code="ServiceUnavailable", message="try later")

        monkeypatch.setattr(
            "ms365_intent_mcp.intent.resolve.impl.compose_resolve",
            _fake,
        )

        payload = ResolvePayload.model_validate({
            "url": "https://teams.microsoft.com/l/chat/19:abc123def456@thread.v2/0"
        })
        response = await _resolve_v1_impl(ctx, payload)

        assert isinstance(response, ErrorResponse)
        assert response.type == "error"
        assert response.code == "graph_api_error"
        assert response.retryable is True
        assert "ServiceUnavailable" in response.message
