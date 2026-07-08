"""Unit tests for _my_day_v1_impl — mocked context, no FastMCP."""

from datetime import date
from unittest.mock import AsyncMock, MagicMock

import pytest

from ms365_intent_mcp.graph import GraphAPIError
from ms365_intent_mcp.intent.my_day.impl import _my_day_v1_impl
from ms365_intent_mcp.intent.my_day.schemas import (
    MailSummary,
    MyDayPayload,
    MyDaySummary,
    TeamsActivitySummary,
)
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


class TestMyDayV1Happy:
    @pytest.mark.asyncio
    async def test_happy_path_returns_my_day_summary(self, monkeypatch):
        ctx, _, _ = _mock_ctx()

        async def _fake(client, permissions, date, timezone):
            return {"date": date, "events": [], "mail": {}, "teams": {}}, "### Calendar\nNo events today.\n\n### Mail\n3 unread."

        monkeypatch.setattr(
            "ms365_intent_mcp.intent.my_day.impl.compose_my_day",
            _fake,
        )

        payload = MyDayPayload(date=date(2026, 7, 8))
        response = await _my_day_v1_impl(ctx, payload)

        assert isinstance(response, MyDaySummary)
        assert response.type == "my_day_summary"
        assert response.date == date(2026, 7, 8)
        assert "Calendar" in response.rendered_markdown
        assert response.events == []
        assert isinstance(response.mail, MailSummary)
        assert isinstance(response.teams, TeamsActivitySummary)

    @pytest.mark.asyncio
    async def test_no_date_defaults_to_today(self, monkeypatch):
        ctx, _, _ = _mock_ctx()
        received_dates = []

        async def _fake(client, permissions, date_str, timezone):
            received_dates.append(date_str)
            return {"date": date_str, "events": [], "mail": {}, "teams": {}}, "### Calendar\nNo events."

        monkeypatch.setattr(
            "ms365_intent_mcp.intent.my_day.impl.compose_my_day",
            _fake,
        )

        payload = MyDayPayload()
        response = await _my_day_v1_impl(ctx, payload)

        assert isinstance(response, MyDaySummary)
        assert len(received_dates) == 1
        assert received_dates[0] == date.today().strftime("%Y-%m-%d")
        assert response.date == date.today()

    @pytest.mark.asyncio
    async def test_explicit_date_is_respected(self, monkeypatch):
        ctx, _, _ = _mock_ctx()
        received_dates = []

        async def _fake(client, permissions, date_str, timezone):
            received_dates.append(date_str)
            return {"date": date_str, "events": [], "mail": {}, "teams": {}}, "### Calendar\n2 events."

        monkeypatch.setattr(
            "ms365_intent_mcp.intent.my_day.impl.compose_my_day",
            _fake,
        )

        target = date(2026, 1, 15)
        payload = MyDayPayload(date=target)
        response = await _my_day_v1_impl(ctx, payload)

        assert response.date == target
        assert received_dates[0] == "2026-01-15"

    @pytest.mark.asyncio
    async def test_stub_fields_are_empty(self, monkeypatch):
        """Structured fields reflect what the composer returns."""
        ctx, _, _ = _mock_ctx()

        async def _fake(client, permissions, date_str, timezone):
            return {"date": date_str, "events": [], "mail": {}, "teams": {}}, "content"

        monkeypatch.setattr(
            "ms365_intent_mcp.intent.my_day.impl.compose_my_day",
            _fake,
        )

        payload = MyDayPayload(date=date(2026, 7, 8))
        response = await _my_day_v1_impl(ctx, payload)

        assert response.events == []
        assert response.mail.unread_count == 0
        assert response.mail.relevant_count == 0
        assert response.mail.flagged_count == 0
        assert response.teams.recent_message_count == 0


class TestMyDayV1Errors:
    @pytest.mark.asyncio
    async def test_graph_api_error_returns_error_response(self, monkeypatch):
        ctx, _, _ = _mock_ctx()

        async def _fake(client, permissions, date_str, timezone):
            raise GraphAPIError(status_code=503, error_code="ServiceUnavailable", message="try later")

        monkeypatch.setattr(
            "ms365_intent_mcp.intent.my_day.impl.compose_my_day",
            _fake,
        )

        payload = MyDayPayload(date=date(2026, 7, 8))
        response = await _my_day_v1_impl(ctx, payload)

        assert isinstance(response, ErrorResponse)
        assert response.type == "error"
        assert response.code == "graph_api_error"
        assert response.retryable is True
        assert "ServiceUnavailable" in response.message
