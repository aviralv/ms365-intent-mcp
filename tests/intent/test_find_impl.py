"""Unit tests for _find_impl — mocked context, no FastMCP."""

from unittest.mock import AsyncMock, MagicMock

import pytest
from pydantic import ValidationError

from ms365_intent_mcp.graph import GraphAPIError
from ms365_intent_mcp.intent._shared import ErrorResponse
from ms365_intent_mcp.intent.find.impl import _find_impl
from ms365_intent_mcp.intent.find.schemas import FindPayload, FindResults


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


class TestFindV1Happy:
    @pytest.mark.asyncio
    async def test_happy_path_returns_find_results(self, monkeypatch):
        ctx, _, _ = _mock_ctx()

        async def _fake(client, permissions, query, search_type):
            return {"query": query, "hits": []}, "### Results\n2 hits found."

        monkeypatch.setattr(
            "ms365_intent_mcp.intent.find.impl.compose_find",
            _fake,
        )

        payload = FindPayload(query="project roadmap")
        response = await _find_impl(ctx, payload)

        assert isinstance(response, FindResults)
        assert response.type == "find_results"
        assert response.query == "project roadmap"
        assert response.hits == []
        assert "Results" in response.rendered_markdown

    @pytest.mark.asyncio
    async def test_entity_type_passed_as_search_type(self, monkeypatch):
        ctx, _, _ = _mock_ctx()
        captured = {}

        async def _fake(client, permissions, query, search_type):
            captured["search_type"] = search_type
            return {"query": query, "hits": []}, "### Results\n1 hit."

        monkeypatch.setattr(
            "ms365_intent_mcp.intent.find.impl.compose_find",
            _fake,
        )

        payload = FindPayload(query="budget", entity_type="email")
        await _find_impl(ctx, payload)

        assert captured["search_type"] == "email"

    @pytest.mark.asyncio
    async def test_no_entity_type_passes_none_as_search_type(self, monkeypatch):
        ctx, _, _ = _mock_ctx()
        captured = {}

        async def _fake(client, permissions, query, search_type):
            captured["search_type"] = search_type
            return {"query": query, "hits": []}, "### Results\n3 hits."

        monkeypatch.setattr(
            "ms365_intent_mcp.intent.find.impl.compose_find",
            _fake,
        )

        payload = FindPayload(query="meeting notes")
        await _find_impl(ctx, payload)

        assert captured["search_type"] is None

    def test_invalid_entity_type_rejected_by_literal(self):
        with pytest.raises(ValidationError) as exc_info:
            FindPayload(query="test", entity_type="audio")
        assert "entity_type" in str(exc_info.value) or "audio" in str(exc_info.value)


    @pytest.mark.asyncio
    async def test_invalid_hit_is_dropped_valid_hit_passes_through(self, monkeypatch, caplog):
        """Malformed hits should be dropped with a warning; valid hits remain."""
        import logging
        ctx, _, _ = _mock_ctx()

        async def _fake(client, permissions, query, search_type):
            return {
                "query": query,
                "hits": [
                    # valid email hit — all required fields present
                    {"kind": "email", "subject": "Budget Q3", "sender": "finance@co.com", "body_preview": "See attached"},
                    # invalid email hit — missing required fields (subject, sender, body_preview)
                    {"kind": "email", "totally": "wrong"},
                ],
            }, "### Results"

        monkeypatch.setattr(
            "ms365_intent_mcp.intent.find.impl.compose_find",
            _fake,
        )

        with caplog.at_level(logging.WARNING, logger="ms365_intent_mcp"):
            payload = FindPayload(query="budget")
            response = await _find_impl(ctx, payload)

        assert isinstance(response, FindResults)
        assert len(response.hits) == 1
        assert response.hits[0].subject == "Budget Q3"
        assert any("malformed" in r.message for r in caplog.records)


    @pytest.mark.asyncio
    async def test_graph_api_error_returns_error_response(self, monkeypatch):
        ctx, _, _ = _mock_ctx()

        async def _fake(client, permissions, query, search_type):
            raise GraphAPIError(
                status_code=503,
                error_code="ServiceUnavailable",
                message="try later",
            )

        monkeypatch.setattr(
            "ms365_intent_mcp.intent.find.impl.compose_find",
            _fake,
        )

        payload = FindPayload(query="test")
        response = await _find_impl(ctx, payload)

        assert isinstance(response, ErrorResponse)
        assert response.type == "error"
        assert response.code == "rate_limited"
        assert response.retryable is True
        assert "ServiceUnavailable" in response.message
