"""Tests for _extract_event_links — absolute-only, deduped, joinUrl-filtered."""

from ms365_intent_mcp.formatters import _extract_event_links


def _event(join_url: str | None = None) -> dict:
    return {"onlineMeeting": {"joinUrl": join_url}} if join_url else {}


def test_extracts_absolute_links():
    html = '<a href="https://wiki.example.com/x">Spec</a> and <a href="https://jira.example.com/AB-1">Ticket</a>'
    assert _extract_event_links(html, _event()) == [
        "https://wiki.example.com/x",
        "https://jira.example.com/AB-1",
    ]


def test_drops_relative_links():
    html = '<a href="/docs/123">rel</a> <a href="https://ok.example.com">abs</a>'
    assert _extract_event_links(html, _event()) == ["https://ok.example.com"]


def test_removes_event_join_url():
    join = "https://teams.microsoft.com/l/meetup-join/19:abc@thread.v2/0"
    html = f'<a href="{join}">Join</a> <a href="https://wiki.example.com/x">Spec</a>'
    assert _extract_event_links(html, _event(join)) == ["https://wiki.example.com/x"]


def test_dedupes():
    html = '<a href="https://a.example.com">1</a> <a href="https://a.example.com">2</a>'
    assert _extract_event_links(html, _event()) == ["https://a.example.com"]


def test_empty_body_returns_empty():
    assert _extract_event_links("", _event()) == []
