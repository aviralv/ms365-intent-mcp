"""Unit tests for _people_v1_impl — mocked context, no FastMCP."""

from unittest.mock import AsyncMock, MagicMock

import pytest

from ms365_intent_mcp.graph import GraphAPIError
from ms365_intent_mcp.intent._shared import ErrorResponse
from ms365_intent_mcp.intent.people.impl import _people_v1_impl
from ms365_intent_mcp.intent.people.schemas import PeoplePayload, PersonDetail


def _mock_ctx():
    """Build a mocked FastMCP Context with the three deps the impls need."""
    ctx = MagicMock()
    config = MagicMock()
    client = AsyncMock()
    permissions = MagicMock()
    ctx.request_context.lifespan_context = {
        "config": config,
        "client": client,
        "permissions": permissions,
    }
    return ctx, client, permissions


class TestPeopleV1Impl:
    @pytest.mark.asyncio
    async def test_happy_path_returns_person_detail(self, monkeypatch):
        """Happy path: compose_people returns (data, markdown) → PersonDetail returned."""
        ctx, _, _ = _mock_ctx()
        markdown = "## Avi Vaid\n**Title:** Senior PM"

        async def _fake(client_arg, perms_arg, query):
            return {"name": "Avi Vaid", "email": "avi@example.com", "job_title": "Senior PM", "recent_mail": [], "recent_chat": None}, markdown

        monkeypatch.setattr(
            "ms365_intent_mcp.intent.people.impl.compose_people",
            _fake,
        )

        payload = PeoplePayload(query="Avi Vaid")
        response = await _people_v1_impl(ctx, payload)

        assert isinstance(response, PersonDetail)
        assert response.type == "person_detail"
        assert response.name == "Avi Vaid"
        assert response.rendered_markdown == markdown
        assert response.recent_mail == []

    @pytest.mark.asyncio
    async def test_query_passed_to_composer(self, monkeypatch):
        """The query from the payload is forwarded to compose_people."""
        ctx, _, _ = _mock_ctx()
        captured = []

        async def _fake(client_arg, perms_arg, query):
            captured.append(query)
            return {"name": query, "email": "", "job_title": None, "recent_mail": [], "recent_chat": None}, "markdown"

        monkeypatch.setattr(
            "ms365_intent_mcp.intent.people.impl.compose_people",
            _fake,
        )

        payload = PeoplePayload(query="john.doe@example.com")
        await _people_v1_impl(ctx, payload)

        assert captured == ["john.doe@example.com"]

    @pytest.mark.asyncio
    async def test_graph_api_error_returns_error_response(self, monkeypatch):
        """GraphAPIError from compose_people → ErrorResponse with graph_api_error code."""
        ctx, _, _ = _mock_ctx()

        async def _fake(client_arg, perms_arg, query):
            raise GraphAPIError(status_code=500, error_code="InternalError", message="server broke")

        monkeypatch.setattr(
            "ms365_intent_mcp.intent.people.impl.compose_people",
            _fake,
        )

        payload = PeoplePayload(query="anyone")
        response = await _people_v1_impl(ctx, payload)

        assert isinstance(response, ErrorResponse)
        assert response.type == "error"
        assert response.code == "graph_api_error"
        assert "InternalError" in response.message

    @pytest.mark.asyncio
    async def test_graph_api_429_is_retryable(self, monkeypatch):
        """429 from Graph → ErrorResponse with retryable=True."""
        ctx, _, _ = _mock_ctx()

        async def _fake(client_arg, perms_arg, query):
            raise GraphAPIError(status_code=429, error_code="TooManyRequests", message="slow down")

        monkeypatch.setattr(
            "ms365_intent_mcp.intent.people.impl.compose_people",
            _fake,
        )

        payload = PeoplePayload(query="someone")
        response = await _people_v1_impl(ctx, payload)

        assert isinstance(response, ErrorResponse)
        assert response.retryable is True


    @pytest.mark.asyncio
    async def test_partial_composer_dict_no_key_error(self, monkeypatch):
        """Partial composer response (missing name/recent_mail) must not raise KeyError."""
        ctx, _, _ = _mock_ctx()

        async def _fake(client_arg, perms_arg, query):
            # Intentionally missing 'name' and 'recent_mail' keys
            return {"email": "partial@example.com"}, "partial markdown"

        monkeypatch.setattr(
            "ms365_intent_mcp.intent.people.impl.compose_people",
            _fake,
        )

        payload = PeoplePayload(query="partial")
        response = await _people_v1_impl(ctx, payload)

        assert isinstance(response, PersonDetail)
        assert response.name == ""
        assert response.recent_mail == []
        assert response.email == "partial@example.com"


    def test_empty_query_rejected(self):
        """min_length=1 on query: empty string must raise ValidationError."""
        import pytest
        from pydantic import ValidationError

        with pytest.raises(ValidationError):
            PeoplePayload(query="")

    def test_valid_query_accepted(self):
        payload = PeoplePayload(query="Avi")
        assert payload.query == "Avi"

    def test_extra_fields_rejected(self):
        """extra='forbid' on PeoplePayload: unknown fields must raise ValidationError."""
        import pytest
        from pydantic import ValidationError

        with pytest.raises(ValidationError):
            PeoplePayload.model_validate({"query": "Avi", "unexpected": "field"})
