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


def test_filters_teams_boilerplate_links():
    """Teams-invite boilerplate URLs (help, dial-in, meeting-options, meet,
    meetup-join) are noise — only the substantive link survives."""
    html = (
        '<a href="https://aka.ms/JoinTeamsMeeting?omkt=en-US">Need help?</a>'
        '<a href="https://dialin.teams.microsoft.com/ba77?id=1">Find a local number</a>'
        '<a href="https://teams.microsoft.com/meetingOptions/?organizerId=x">Meeting options</a>'
        '<a href="https://teams.microsoft.com/meet/369947989999958?p=abc">Join</a>'
        '<a href="https://teams.microsoft.com/l/meetup-join/19:abc@thread.v2/0">System reference</a>'
        '<a href="https://leanix.atlassian.net/wiki/spaces/NOVA/pages/9772893179/Knowledge+Cafe">Agenda</a>'
    )
    assert _extract_event_links(html, _event()) == [
        "https://leanix.atlassian.net/wiki/spaces/NOVA/pages/9772893179/Knowledge+Cafe",
    ]


def test_decodes_html_entities_in_url():
    """Graph bodies HTML-escape & as &amp; inside href — decode so the URL works."""
    html = '<a href="https://x.example.com/p?a=1&amp;b=2">link</a>'
    assert _extract_event_links(html, _event()) == ["https://x.example.com/p?a=1&b=2"]


def test_no_agenda_meeting_yields_empty_links():
    """A pure-boilerplate Teams invite (no real content) → no links at all."""
    html = (
        '<a href="https://aka.ms/JoinTeamsMeeting?omkt=en-US">Need help?</a>'
        '<a href="https://teams.microsoft.com/meet/123?p=x">Join</a>'
        '<a href="https://dialin.teams.cloud.microsoft/ba77?id=2">Find a local number</a>'
    )
    assert _extract_event_links(html, _event()) == []
