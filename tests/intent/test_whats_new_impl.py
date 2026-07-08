"""Unit tests for _whats_new_impl — mocked context, no FastMCP."""

from datetime import datetime, timezone
from unittest.mock import AsyncMock, MagicMock

import pytest

from ms365_intent_mcp.graph import GraphAPIError
from ms365_intent_mcp.intent.whats_new.impl import _whats_new_impl
from ms365_intent_mcp.intent.whats_new.schemas import WhatsNewPayload, WhatsNewSummary
from ms365_intent_mcp.intent._shared import ErrorResponse


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


class TestWhatsNewV1Happy:
    @pytest.mark.asyncio
    async def test_happy_path_returns_whats_new_summary(self, monkeypatch):
        ctx, _, _ = _mock_ctx()
        since = datetime(2026, 7, 1, 9, 0, 0, tzinfo=timezone.utc)

        async def _fake(client, permissions, since, scope, timezone):
            return {"since": since, "mail": [], "events": [], "teams": []}, "### Calendar\nTeam sync.\n\n### Mail\n2 unread."

        monkeypatch.setattr(
            "ms365_intent_mcp.intent.whats_new.impl.compose_whats_new",
            _fake,
        )

        payload = WhatsNewPayload(since=since, scope="all")
        response = await _whats_new_impl(ctx, payload)

        assert isinstance(response, WhatsNewSummary)
        assert response.type == "whats_new_summary"
        assert response.since == since
        assert "Calendar" in response.rendered_markdown
        assert response.mail == []
        assert response.events == []
        assert response.teams == []

    @pytest.mark.asyncio
    async def test_scope_all_translates_to_none(self, monkeypatch):
        """scope='all' must be translated to None before calling the composer."""
        ctx, _, _ = _mock_ctx()
        since = datetime(2026, 7, 1, 0, 0, 0, tzinfo=timezone.utc)
        captured_kwargs = {}

        async def _fake(client, permissions, since, scope, timezone):
            captured_kwargs["scope"] = scope
            return {"since": since, "mail": [], "events": [], "teams": []}, "Nothing new."

        monkeypatch.setattr(
            "ms365_intent_mcp.intent.whats_new.impl.compose_whats_new",
            _fake,
        )

        payload = WhatsNewPayload(since=since, scope="all")
        await _whats_new_impl(ctx, payload)

        assert captured_kwargs["scope"] is None

    @pytest.mark.asyncio
    async def test_scope_teams_passes_through(self, monkeypatch):
        """scope='teams' must be forwarded to the composer unchanged."""
        ctx, _, _ = _mock_ctx()
        since = datetime(2026, 7, 1, 0, 0, 0, tzinfo=timezone.utc)
        captured_kwargs = {}

        async def _fake(client, permissions, since, scope, timezone):
            captured_kwargs["scope"] = scope
            return {"since": since, "mail": [], "events": [], "teams": []}, "### Teams\nNo new messages."

        monkeypatch.setattr(
            "ms365_intent_mcp.intent.whats_new.impl.compose_whats_new",
            _fake,
        )

        payload = WhatsNewPayload(since=since, scope="teams")
        await _whats_new_impl(ctx, payload)

        assert captured_kwargs["scope"] == "teams"

    @pytest.mark.asyncio
    async def test_stub_fields_are_empty(self, monkeypatch):
        """Structured fields reflect what the composer returns."""
        ctx, _, _ = _mock_ctx()
        since = datetime(2026, 7, 1, 0, 0, 0, tzinfo=timezone.utc)

        async def _fake(client, permissions, since, scope, timezone):
            return {"since": since, "mail": [], "events": [], "teams": []}, "content"

        monkeypatch.setattr(
            "ms365_intent_mcp.intent.whats_new.impl.compose_whats_new",
            _fake,
        )

        payload = WhatsNewPayload(since=since)
        response = await _whats_new_impl(ctx, payload)

        assert response.mail == []
        assert response.events == []
        assert response.teams == []


class TestWhatsNewV1Errors:
    @pytest.mark.asyncio
    async def test_graph_api_error_returns_error_response(self, monkeypatch):
        ctx, _, _ = _mock_ctx()
        since = datetime(2026, 7, 1, 0, 0, 0, tzinfo=timezone.utc)

        async def _fake(client, permissions, since, scope, timezone):
            raise GraphAPIError(status_code=429, error_code="TooManyRequests", message="slow down")

        monkeypatch.setattr(
            "ms365_intent_mcp.intent.whats_new.impl.compose_whats_new",
            _fake,
        )

        payload = WhatsNewPayload(since=since)
        response = await _whats_new_impl(ctx, payload)

        assert isinstance(response, ErrorResponse)
        assert response.type == "error"
        assert response.code == "rate_limited"
        assert response.retryable is True
        assert "TooManyRequests" in response.message
