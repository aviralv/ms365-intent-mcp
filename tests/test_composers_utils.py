"""Tests for shared composer utilities."""

from ms365_intent_mcp.composers._utils import _error_reason, _escape_odata, _build_mail_summary, NOISE_PATTERNS
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


class TestEscapeOdata:
    def test_no_quotes_unchanged(self):
        assert _escape_odata("hello world") == "hello world"

    def test_single_quote_doubled(self):
        assert _escape_odata("O'Brien") == "O''Brien"

    def test_multiple_quotes(self):
        assert _escape_odata("it's a 'test'") == "it''s a ''test''"

    def test_empty_string(self):
        assert _escape_odata("") == ""


class TestBuildMailSummary:
    def test_filters_noise(self):
        msgs = [
            {"from": {"emailAddress": {"address": "noreply@github.com", "name": "GitHub"}}, "subject": "PR", "importance": "normal"},
            {"from": {"emailAddress": {"address": "alice@company.com", "name": "Alice"}}, "subject": "Important", "importance": "high"},
        ]
        summary = _build_mail_summary(msgs)
        assert summary["all_count"] == 2
        assert summary["relevant_count"] == 1
        assert len(summary["high_importance"]) == 1
        assert summary["high_importance"][0]["subject"] == "Important"
        assert "github.com" in NOISE_PATTERNS


def test_chat_enumeration_helpers_live_in_utils():
    """Task 1 move: both helpers are importable from _utils (shared boundary)."""
    from ms365_intent_mcp.composers._utils import (
        _list_user_chats,
        _prefilter_chats_by_query,
    )
    assert callable(_list_user_chats)
    assert callable(_prefilter_chats_by_query)


def test_prefilter_narrows_by_member_name():
    from ms365_intent_mcp.composers._utils import _prefilter_chats_by_query
    chats = [
        {"id": "1", "members": [{"displayName": "Yevhen Kushnirenko"}]},
        {"id": "2", "members": [{"displayName": "Bob Jones"}]},
    ]
    matched, words = _prefilter_chats_by_query(chats, "Yevhen")
    assert [c["id"] for c in matched] == ["1"]
    assert words == {"yevhen"}
