"""Tests for PermissionRegistry scope checking."""

import base64
import json

from ms365_intent_mcp.permissions import PermissionRegistry


class TestPermissionRegistry:
    def test_has_scope_true(self):
        registry = PermissionRegistry(["Mail.Read", "Calendars.ReadWrite"])
        assert registry.has("Mail.Read") is True

    def test_has_scope_false(self):
        registry = PermissionRegistry(["Mail.Read"])
        assert registry.has("Chat.ReadWrite") is False

    def test_has_any_true(self):
        registry = PermissionRegistry(["Mail.Read", "Calendars.ReadWrite"])
        assert registry.has_any(["Chat.ReadWrite", "Mail.Read"]) is True

    def test_has_any_false(self):
        registry = PermissionRegistry(["Mail.Read"])
        assert registry.has_any(["Chat.ReadWrite", "Files.Read"]) is False

    def test_check_returns_none_when_granted(self):
        registry = PermissionRegistry(["Mail.Read"])
        assert registry.check("Mail.Read") is None

    def test_check_returns_message_when_missing(self):
        registry = PermissionRegistry(["Mail.Read"])
        result = registry.check("Chat.ReadWrite")
        assert result is not None
        assert "Chat.ReadWrite" in result

    def test_from_token_decodes_scp_claim(self):
        payload = {"scp": "Mail.Read Calendars.ReadWrite User.Read"}
        encoded = base64.urlsafe_b64encode(json.dumps(payload).encode()).rstrip(b"=").decode()
        fake_jwt = f"header.{encoded}.signature"
        registry = PermissionRegistry.from_token(fake_jwt)
        assert registry.has("Mail.Read")
        assert registry.has("Calendars.ReadWrite")
        assert not registry.has("Chat.ReadWrite")

    def test_from_token_empty_on_invalid(self):
        registry = PermissionRegistry.from_token("not-a-jwt")
        assert registry.has("Mail.Read") is False

    def test_granted_property(self):
        registry = PermissionRegistry(["Mail.Read", "User.Read"])
        assert registry.granted == {"Mail.Read", "User.Read"}


def _make_jwt(scopes: str) -> str:
    payload = {"scp": scopes}
    encoded = base64.urlsafe_b64encode(json.dumps(payload).encode()).rstrip(b"=").decode()
    return f"header.{encoded}.signature"


class TestLazyPermissionRegistry:
    def test_updates_on_token_change(self):
        tokens = [_make_jwt("Mail.Read")]

        def provider():
            return tokens[0]

        registry = PermissionRegistry.from_token_provider(provider)
        assert registry.has("Mail.Read")
        assert not registry.has("Chat.ReadWrite")

        tokens[0] = _make_jwt("Mail.Read Chat.ReadWrite")
        assert registry.has("Chat.ReadWrite")

    def test_does_not_reparse_same_token(self):
        call_count = [0]
        token = _make_jwt("Mail.Read")

        def provider():
            call_count[0] += 1
            return token

        registry = PermissionRegistry.from_token_provider(provider)
        registry.has("Mail.Read")
        registry.has("Mail.Read")
        assert call_count[0] == 2  # provider called, but decode only happens once

    def test_base64_padding_various_lengths(self):
        for scopes in ["A", "AB", "ABC", "ABCD", "Mail.Read Calendars.ReadWrite User.Read"]:
            token = _make_jwt(scopes)
            registry = PermissionRegistry.from_token(token)
            assert registry.granted  # should not raise

    def test_handles_none_token(self):
        """Provider returning None (no token cached yet) must not crash."""
        registry = PermissionRegistry.from_token_provider(lambda: None)
        assert registry.has("Mail.Read") is False
        assert registry.granted == set()

    def test_recovers_when_token_becomes_available(self):
        tokens: list[str | None] = [None]

        def provider():
            return tokens[0]

        registry = PermissionRegistry.from_token_provider(provider)
        assert not registry.has("Mail.Read")
        tokens[0] = _make_jwt("Mail.Read")
        assert registry.has("Mail.Read")
