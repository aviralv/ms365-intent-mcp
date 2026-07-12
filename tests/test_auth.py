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


class TestSharePointTokenHostGuard:
    """get_sharepoint_token interpolates host into the token scope, so it must
    refuse non-SharePoint hosts even before touching the network (issue #29
    review — defense-in-depth against a caller that skips VroomClient's guard)."""

    def test_rejects_non_sharepoint_host(self, tmp_path):
        config = Config(token_path=tmp_path / "token.json")
        mgr = TokenManager(config)
        with pytest.raises(ValueError, match="non-SharePoint host"):
            mgr.get_sharepoint_token("evil.com")

    def test_rejects_lookalike_host(self, tmp_path):
        config = Config(token_path=tmp_path / "token.json")
        mgr = TokenManager(config)
        with pytest.raises(ValueError, match="non-SharePoint host"):
            mgr.get_sharepoint_token("evilsharepoint.com")

    def test_rejects_empty_host(self, tmp_path):
        config = Config(token_path=tmp_path / "token.json")
        mgr = TokenManager(config)
        with pytest.raises(ValueError):
            mgr.get_sharepoint_token("")

    def test_returns_cached_spo_token_without_network(self, tmp_path):
        config = Config(token_path=tmp_path / "token.json")
        mgr = TokenManager(config)
        mgr._spo_tokens["sap-my.sharepoint.com"] = ("spo-cached", time.time() + 3600)
        assert mgr.get_sharepoint_token("sap-my.sharepoint.com") == "spo-cached"
