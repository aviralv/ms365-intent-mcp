"""Tests for legacy tool deprecation warning rate-limiting in server.py."""

import logging

import pytest

import ms365_intent_mcp.server as server_module


@pytest.fixture(autouse=True)
def _reset_warn_state():
    """Ensure the last-warn dict is empty before each test."""
    server_module._last_legacy_warn.clear()
    yield
    server_module._last_legacy_warn.clear()


class TestWarnLegacyOnce:
    def test_first_call_emits_warning(self, caplog):
        with caplog.at_level(logging.WARNING, logger="ms365_intent_mcp"):
            server_module._warn_legacy_once("my_day")
        assert any("my_day" in r.message for r in caplog.records)

    def test_second_call_within_interval_does_not_emit(self, caplog):
        """Two rapid calls to the same tool should emit only one warning."""
        with caplog.at_level(logging.WARNING, logger="ms365_intent_mcp"):
            server_module._warn_legacy_once("my_day")
            server_module._warn_legacy_once("my_day")
        matching = [r for r in caplog.records if "my_day" in r.message]
        assert len(matching) == 1, f"expected exactly 1 warning, got {len(matching)}"

    def test_different_tools_each_emit_once(self, caplog):
        """Different tools should each get their own warning on first call."""
        with caplog.at_level(logging.WARNING, logger="ms365_intent_mcp"):
            server_module._warn_legacy_once("my_day")
            server_module._warn_legacy_once("find")
        my_day_warns = [r for r in caplog.records if "my_day" in r.message]
        find_warns = [r for r in caplog.records if "'find'" in r.message]
        assert len(my_day_warns) == 1
        assert len(find_warns) == 1

    def test_after_interval_expiry_emits_again(self, caplog):
        """After the interval expires, the next call should re-emit."""
        original_interval = server_module._LEGACY_WARN_INTERVAL_SECONDS
        server_module._LEGACY_WARN_INTERVAL_SECONDS = -1  # always expired
        try:
            with caplog.at_level(logging.WARNING, logger="ms365_intent_mcp"):
                server_module._warn_legacy_once("my_day")
                server_module._warn_legacy_once("my_day")
            matching = [r for r in caplog.records if "my_day" in r.message]
            assert len(matching) == 2, f"expected 2 warnings after interval reset, got {len(matching)}"
        finally:
            server_module._LEGACY_WARN_INTERVAL_SECONDS = original_interval
