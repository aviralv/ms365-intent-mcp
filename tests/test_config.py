"""Tests for Config defaults and env overrides."""

import os
from pathlib import Path
from unittest.mock import patch

from ms365_intent_mcp.config import Config


class TestConfigDefaults:
    def test_default_token_path(self):
        config = Config()
        assert config.token_path == Path.home() / ".config" / "ms365-intent-mcp" / "token.json"

    def test_default_timezone(self):
        config = Config()
        assert config.default_timezone == "Europe/Berlin"

    def test_default_graph_url(self):
        config = Config()
        assert config.graph_base_url == "https://graph.microsoft.com/v1.0"

    def test_scopes_includes_calendar(self):
        config = Config()
        assert "Calendars.ReadWrite" in config.scopes

    def test_scopes_includes_people_read(self):
        config = Config()
        assert "People.Read" in config.scopes


class TestConfigEnvOverride:
    def test_timezone_override(self):
        with patch.dict(os.environ, {"MS365_INTENT_DEFAULT_TIMEZONE": "America/New_York"}):
            config = Config()
            assert config.default_timezone == "America/New_York"

    def test_token_path_override(self):
        with patch.dict(os.environ, {"MS365_INTENT_TOKEN_PATH": "/tmp/test-token.json"}):
            config = Config()
            assert config.token_path == Path("/tmp/test-token.json")
