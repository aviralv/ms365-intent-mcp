"""Unit tests for _find_v1_impl — mocked context, no FastMCP."""

from unittest.mock import AsyncMock, MagicMock

import pytest
from pydantic import ValidationError

from ms365_intent_mcp.graph import GraphAPIError
from ms365_intent_mcp.intent._shared import ErrorResponse
from ms365_intent_mcp.intent.find.impl import _find_v1_impl
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
        response = await _find_v1_impl(ctx, payload)

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
        await _find_v1_impl(ctx, payload)

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
        await _find_v1_impl(ctx, payload)

        assert captured["search_type"] is None

    def test_invalid_entity_type_rejected_by_literal(self):
        with pytest.raises(ValidationError) as exc_info:
            FindPayload(query="test", entity_type="audio")
        assert "entity_type" in str(exc_info.value) or "audio" in str(exc_info.value)


class TestFindV1Errors:
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
        response = await _find_v1_impl(ctx, payload)

        assert isinstance(response, ErrorResponse)
        assert response.type == "error"
        assert response.code == "graph_api_error"
        assert response.retryable is True
        assert "ServiceUnavailable" in response.message
