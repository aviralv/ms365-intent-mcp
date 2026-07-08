"""Unit tests for _compose_v1_impl — mocked context, no FastMCP."""

from unittest.mock import AsyncMock, MagicMock

import pytest

from ms365_intent_mcp.intent._helpers import idempotency_clear
from ms365_intent_mcp.intent.compose.impl import _compose_v1_impl
from ms365_intent_mcp.intent.compose.schemas import (
    ComposeEmail,
    ComposeEvent,
    ComposeTeamsMessage,
    EmailDraftCreated,
    EventCreated,
    TeamsMessageSent,
)


@pytest.fixture(autouse=True)
def _clear_cache():
    """Ensure idempotency cache doesn't leak state across tests."""
    idempotency_clear()
    yield
    idempotency_clear()


def _mock_ctx():
    """Build a mocked FastMCP Context with the three deps the impls need."""
    ctx = MagicMock()
    config = MagicMock(default_timezone="Europe/Berlin")
    client = AsyncMock()
    permissions = MagicMock()
    permissions.check = MagicMock(return_value=None)  # scope always OK
    ctx.request_context.lifespan_context = {
        "config": config,
        "client": client,
        "permissions": permissions,
    }
    return ctx, client, permissions


class TestComposeV1Email:
    @pytest.mark.asyncio
    async def test_new_email_returns_typed_response(self, monkeypatch):
        ctx, _, _ = _mock_ctx()

        async def _fake(client_arg, perms_arg, action_type, params):
            return "✅ Draft created\n**Subject:** Hi"

        monkeypatch.setattr(
            "ms365_intent_mcp.intent.compose.impl.compose_action",
            _fake,
        )

        payload = ComposeEmail.model_validate({
            "type": "email",
            "mode": "new",
            "to": [{"email": "a@b.com"}],
            "subject": "Hi",
            "body": "Hello",
        })
        response = await _compose_v1_impl(ctx, payload)

        assert isinstance(response, EmailDraftCreated)
        assert response.type == "email_draft_created"
        assert "Draft created" in response.rendered_markdown

    @pytest.mark.asyncio
    async def test_reply_dispatches_to_reply_draft(self, monkeypatch):
        ctx, _, _ = _mock_ctx()
        captured_action = []

        async def _fake(client_arg, perms_arg, action_type, params):
            captured_action.append(action_type)
            return "✅ Reply draft created"

        monkeypatch.setattr(
            "ms365_intent_mcp.intent.compose.impl.compose_action",
            _fake,
        )

        payload = ComposeEmail.model_validate({
            "type": "email",
            "mode": "reply",
            "in_reply_to_message_id": "AAM123",
            "body": "Reply body",
        })
        await _compose_v1_impl(ctx, payload)

        assert len(captured_action) == 1
        assert captured_action[0].value == "reply_draft"


class TestComposeV1Event:
    @pytest.mark.asyncio
    async def test_event_returns_typed_response(self, monkeypatch):
        ctx, _, _ = _mock_ctx()

        async def _fake(client_arg, perms_arg, action_type, params):
            return "✅ Event created\n**Subject:** Sync"

        monkeypatch.setattr(
            "ms365_intent_mcp.intent.compose.impl.compose_action",
            _fake,
        )

        payload = ComposeEvent.model_validate({
            "type": "event",
            "subject": "Sync",
            "start": "2026-07-08T10:00:00Z",
            "end": "2026-07-08T11:00:00Z",
            "timezone": "Europe/Berlin",
        })
        response = await _compose_v1_impl(ctx, payload)

        assert isinstance(response, EventCreated)
        assert response.subject == "Sync"
        assert "Event created" in response.rendered_markdown


class TestComposeV1TeamsMessage:
    @pytest.mark.asyncio
    async def test_teams_message_returns_typed_response(self, monkeypatch):
        ctx, _, _ = _mock_ctx()

        async def _fake(client_arg, perms_arg, action_type, params):
            return "✅ Message sent to Teams chat."

        monkeypatch.setattr(
            "ms365_intent_mcp.intent.compose.impl.compose_action",
            _fake,
        )

        payload = ComposeTeamsMessage.model_validate({
            "type": "teams_message",
            "chat_id": "19:abc@thread.v2",
            "content": "hi",
        })
        response = await _compose_v1_impl(ctx, payload)

        assert isinstance(response, TeamsMessageSent)
        assert response.chat_id == "19:abc@thread.v2"


class TestIdempotencyKey:
    @pytest.mark.asyncio
    async def test_second_call_returns_cached(self, monkeypatch):
        ctx, _, _ = _mock_ctx()
        call_count = 0

        async def _fake(client_arg, perms_arg, action_type, params):
            nonlocal call_count
            call_count += 1
            return f"✅ Draft created (call {call_count})"

        monkeypatch.setattr(
            "ms365_intent_mcp.intent.compose.impl.compose_action",
            _fake,
        )

        payload = ComposeEmail.model_validate({
            "type": "email",
            "mode": "new",
            "to": [{"email": "a@b.com"}],
            "subject": "Hi",
            "body": "Hello",
            "idempotency_key": "abc-123",
        })
        r1 = await _compose_v1_impl(ctx, payload)
        r2 = await _compose_v1_impl(ctx, payload)

        assert call_count == 1, "second call should be served from cache"
        assert r1.rendered_markdown == r2.rendered_markdown

    @pytest.mark.asyncio
    async def test_past_ttl_reexecutes(self, monkeypatch):
        """When TTL expires, subsequent calls must re-execute the underlying work."""
        import ms365_intent_mcp.intent._helpers as h

        original_ttl = h._IDEMPOTENCY_TTL_SECONDS
        h._IDEMPOTENCY_TTL_SECONDS = 0  # instant expiry
        try:
            ctx, _, _ = _mock_ctx()
            call_count = 0

            async def _fake(*args, **kwargs):
                nonlocal call_count
                call_count += 1
                return f"call {call_count}"

            monkeypatch.setattr(
                "ms365_intent_mcp.intent.compose.impl.compose_action",
                _fake,
            )

            payload = ComposeEmail.model_validate({
                "type": "email",
                "mode": "new",
                "to": [{"email": "a@b.com"}],
                "subject": "Hi",
                "body": "Hi",
                "idempotency_key": "expiry-test",
            })
            await _compose_v1_impl(ctx, payload)
            await _compose_v1_impl(ctx, payload)
            assert call_count == 2, "past-TTL call should re-execute"
        finally:
            h._IDEMPOTENCY_TTL_SECONDS = original_ttl

    @pytest.mark.asyncio
    async def test_no_key_no_cache(self, monkeypatch):
        """Without idempotency_key, every call executes."""
        ctx, _, _ = _mock_ctx()
        call_count = 0

        async def _fake(*args, **kwargs):
            nonlocal call_count
            call_count += 1
            return f"call {call_count}"

        monkeypatch.setattr(
            "ms365_intent_mcp.intent.compose.impl.compose_action",
            _fake,
        )

        payload = ComposeEmail.model_validate({
            "type": "email",
            "mode": "new",
            "to": [{"email": "a@b.com"}],
            "subject": "Hi",
            "body": "Hi",
        })
        await _compose_v1_impl(ctx, payload)
        await _compose_v1_impl(ctx, payload)
        assert call_count == 2, "no key should mean no cache"
