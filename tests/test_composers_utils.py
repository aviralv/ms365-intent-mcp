"""Tests for shared composer utilities."""

import pytest

from ms365_intent_mcp.composers._utils import _error_reason, NOISE_PATTERNS
from ms365_intent_mcp.graph import GraphAPIError


class TestErrorReason:
    def test_429_returns_rate_limit_message(self):
        err = GraphAPIError(429, "TooManyRequests", "slow down")
        assert _error_reason(err) == "rate limited — retry shortly"

    def test_500_returns_service_error(self):
        err = GraphAPIError(500, "InternalError", "exploded")
        assert _error_reason(err) == "Microsoft service error"

    def test_400_returns_graph_message(self):
        err = GraphAPIError(400, "BadRequest", "invalid param")
        assert _error_reason(err) == "invalid param"

    def test_timeout_string(self):
        assert _error_reason(TimeoutError("timed out")) == "timed out"

    def test_generic_exception_truncated(self):
        msg = "x" * 200
        result = _error_reason(ValueError(msg))
        assert len(result) <= 100

    def test_noise_patterns_include_expected(self):
        assert "noreply@" in NOISE_PATTERNS
        assert "github.com" in NOISE_PATTERNS
