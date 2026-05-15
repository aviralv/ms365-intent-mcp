"""Tests for TokenManager refresh logic."""

import json
import time

import pytest

from ms365_intent_mcp.auth import TokenManager, AuthenticationError
from ms365_intent_mcp.config import Config


class TestTokenManagerCached:
    def test_returns_cached_token_when_not_expired(self, tmp_path):
        config = Config(token_path=tmp_path / "token.json")
        mgr = TokenManager(config)
        mgr._access_token = "cached-token"
        mgr._expires_at = time.time() + 3600
        assert mgr.get_access_token() == "cached-token"


class TestTokenManagerRefresh:
    def test_raises_when_no_token_file(self, tmp_path):
        config = Config(token_path=tmp_path / "nonexistent.json")
        mgr = TokenManager(config)
        with pytest.raises(AuthenticationError):
            mgr.ensure_authenticated()

    def test_raises_when_no_refresh_token(self, tmp_path):
        token_file = tmp_path / "token.json"
        token_file.write_text(json.dumps({"access_token": "old"}))
        config = Config(token_path=token_file)
        mgr = TokenManager(config)
        with pytest.raises(AuthenticationError):
            mgr.ensure_authenticated()
