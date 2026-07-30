"""Unit tests for _compose_impl — mocked context, no FastMCP."""

from unittest.mock import AsyncMock, MagicMock

import pytest

from ms365_intent_mcp.intent._helpers import idempotency_clear, idempotency_lookup, idempotency_store
from ms365_intent_mcp.intent.compose.impl import _compose_impl, _handle_email
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
            return {"draft_id": "draft-abc", "subject": "Hi", "to": [{"email": "a@b.com", "name": "A"}], "web_link": "https://outlook.office.com/mail/inbox"}, "✅ Draft created\n**Subject:** Hi"

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
        response = await _compose_impl(ctx, payload)

        assert isinstance(response, EmailDraftCreated)
        assert response.type == "email_draft_created"
        assert "Draft created" in response.rendered_markdown

    @pytest.mark.asyncio
    async def test_reply_dispatches_to_reply_draft(self, monkeypatch):
        ctx, _, _ = _mock_ctx()
        captured_action = []

        async def _fake(client_arg, perms_arg, action_type, params):
            captured_action.append(action_type)
            return {"draft_id": "reply-1", "subject": "Re: Hi", "to": [], "web_link": "https://outlook.office.com/mail/inbox"}, "✅ Reply draft created"

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
        await _compose_impl(ctx, payload)

        assert len(captured_action) == 1
        assert captured_action[0].value == "reply_draft"


class TestComposeV1Event:
    @pytest.mark.asyncio
    async def test_event_returns_typed_response(self, monkeypatch):
        ctx, _, _ = _mock_ctx()
        captured_params = {}

        async def _fake(client_arg, perms_arg, action_type, params):
            captured_params.update(params)
            return {"event_id": "evt-123", "subject": "Sync", "start": "2026-07-08T10:00:00Z", "end": "2026-07-08T11:00:00Z", "join_url": None}, "✅ Event created\n**Subject:** Sync"

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
        response = await _compose_impl(ctx, payload)

        assert isinstance(response, EventCreated)
        assert response.subject == "Sync"
        assert "Event created" in response.rendered_markdown
        # The timezone passed to the composer must be payload.timezone, not the config default
        assert captured_params["timezone"] == "Europe/Berlin"


class TestComposeV1TeamsMessage:
    @pytest.mark.asyncio
    async def test_teams_message_returns_typed_response(self, monkeypatch):
        ctx, _, _ = _mock_ctx()

        async def _fake(client_arg, perms_arg, action_type, params):
            return {"message_id": "msg-456", "chat_id": "19:abc@thread.v2"}, "✅ Message sent to Teams chat."

        monkeypatch.setattr(
            "ms365_intent_mcp.intent.compose.impl.compose_action",
            _fake,
        )

        payload = ComposeTeamsMessage.model_validate({
            "type": "teams_message",
            "chat_id": "19:abc@thread.v2",
            "content": "hi",
        })
        response = await _compose_impl(ctx, payload)

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
            return {"draft_id": f"d{call_count}", "subject": "Hi", "to": [{"email": "a@b.com", "name": "A"}], "web_link": "https://outlook.office.com/mail/inbox"}, f"✅ Draft created (call {call_count})"

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
        r1 = await _compose_impl(ctx, payload)
        r2 = await _compose_impl(ctx, payload)

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
                return {"draft_id": f"d{call_count}", "subject": "Hi", "to": [{"email": "a@b.com", "name": "A"}], "web_link": "https://outlook.office.com/mail/inbox"}, f"call {call_count}"

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
            await _compose_impl(ctx, payload)
            await _compose_impl(ctx, payload)
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
            return {"draft_id": f"d{call_count}", "subject": "Hi", "to": [{"email": "a@b.com", "name": "A"}], "web_link": "https://outlook.office.com/mail/inbox"}, f"call {call_count}"

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
        await _compose_impl(ctx, payload)
        await _compose_impl(ctx, payload)
        assert call_count == 2, "no key should mean no cache"


class TestIdempotencyLookupDirect:
    """Direct unit tests for idempotency_lookup — no FastMCP overhead."""

    def setup_method(self):
        idempotency_clear()

    def teardown_method(self):
        idempotency_clear()

    def test_lookup_returns_none_when_empty(self):
        assert idempotency_lookup("tool", "key-1") is None

    def test_lookup_returns_stored_value(self):
        sentinel = object()
        idempotency_store("tool", "key-2", sentinel)
        assert idempotency_lookup("tool", "key-2") is sentinel

    def test_lookup_returns_none_after_ttl_expiry(self):
        import ms365_intent_mcp.intent._helpers as h
        original_ttl = h._IDEMPOTENCY_TTL_SECONDS
        h._IDEMPOTENCY_TTL_SECONDS = -1  # always expired (any elapsed time > -1)
        try:
            idempotency_store("tool", "key-ttl", "cached-response")
            result = idempotency_lookup("tool", "key-ttl")
            assert result is None, "expired entry should return None, not the stale response"
        finally:
            h._IDEMPOTENCY_TTL_SECONDS = original_ttl

    def test_lookup_no_op_on_falsy_key(self):
        assert idempotency_lookup("tool", None) is None
        assert idempotency_lookup("tool", "") is None


class TestWrapErrors:
    """Unit tests for wrap_errors — direct function exercise, no FastMCP."""

    @pytest.mark.asyncio
    async def test_429_maps_to_rate_limited(self):
        from ms365_intent_mcp.graph import GraphAPIError
        from ms365_intent_mcp.intent._helpers import wrap_errors
        from ms365_intent_mcp.intent._shared import ErrorResponse

        @wrap_errors("test_tool")
        async def _raises(*args, **kwargs):
            raise GraphAPIError(429, "TooManyRequests", "slow down")

        result = await _raises()
        assert isinstance(result, ErrorResponse)
        assert result.code == "rate_limited"
        assert result.retryable is True

    @pytest.mark.asyncio
    async def test_503_maps_to_rate_limited(self):
        from ms365_intent_mcp.graph import GraphAPIError
        from ms365_intent_mcp.intent._helpers import wrap_errors
        from ms365_intent_mcp.intent._shared import ErrorResponse

        @wrap_errors("test_tool")
        async def _raises(*args, **kwargs):
            raise GraphAPIError(503, "ServiceUnavailable", "try later")

        result = await _raises()
        assert isinstance(result, ErrorResponse)
        assert result.code == "rate_limited"
        assert result.retryable is True

    @pytest.mark.asyncio
    async def test_500_maps_to_graph_api_error(self):
        from ms365_intent_mcp.graph import GraphAPIError
        from ms365_intent_mcp.intent._helpers import wrap_errors
        from ms365_intent_mcp.intent._shared import ErrorResponse

        @wrap_errors("test_tool")
        async def _raises(*args, **kwargs):
            raise GraphAPIError(500, "InternalError", "boom")

        result = await _raises()
        assert isinstance(result, ErrorResponse)
        assert result.code == "graph_api_error"
        assert result.retryable is False

    @pytest.mark.asyncio
    async def test_broad_exception_returns_error_response(self):
        """Unhandled exceptions (e.g. KeyError) must be caught and returned as ErrorResponse."""
        from ms365_intent_mcp.intent._helpers import wrap_errors
        from ms365_intent_mcp.intent._shared import ErrorResponse

        @wrap_errors("test_tool")
        async def _raises(*args, **kwargs):
            raise KeyError("unexpected_key")

        result = await _raises()
        assert isinstance(result, ErrorResponse)
        assert result.code == "graph_api_error"
        assert "KeyError" in result.message
        assert result.retryable is False


class TestHandleEventForward:
    @pytest.mark.asyncio
    async def test_handle_event_forward_returns_event_forwarded(self):
        from ms365_intent_mcp.intent.compose.impl import _handle_event
        from ms365_intent_mcp.intent.compose.schemas import ComposeEvent
        from ms365_intent_mcp.permissions import PermissionRegistry

        client = AsyncMock()
        client.post = AsyncMock(return_value={})
        perms = PermissionRegistry(["Calendars.ReadWrite"])
        payload = ComposeEvent.model_validate({
            "type": "event", "mode": "forward",
            "event_id": "AAMkEVT",
            "to": [{"email": "dana@contoso.com", "name": "Dana Swope"}],
            "comment": "please join",
        })
        resp = await _handle_event(payload, client, perms, "Europe/Berlin")
        assert resp.type == "event_forwarded"
        assert resp.to[0].email == "dana@contoso.com"
        assert client.post.call_args.args[0] == "/me/events/AAMkEVT/forward"


class TestHandleEmailForward:
    @pytest.mark.asyncio
    async def test_handle_email_forward_routes_to_createforward(self):
        from ms365_intent_mcp.permissions import PermissionRegistry

        client = AsyncMock()
        client.post = AsyncMock(return_value={
            "id": "d-1", "subject": "FW: X",
            "toRecipients": [{"emailAddress": {"address": "z@z.com", "name": "Z"}}],
            "webLink": "https://outlook.office.com/mail/x",
        })
        perms = PermissionRegistry(["Mail.ReadWrite"])
        payload = ComposeEmail.model_validate({
            "type": "email", "mode": "forward",
            "in_reply_to_message_id": "m-1",
            "to": [{"email": "z@z.com", "name": "Z"}],
            "body": "fyi",
        })
        resp = await _handle_email(payload, client, perms)
        assert resp.type == "email_draft_created"
        assert client.post.call_args.args[0] == "/me/messages/m-1/createForward"
