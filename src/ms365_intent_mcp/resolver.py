"""URL resolver — regex dispatch table for 7 M365 URL types."""

from __future__ import annotations

import re
import urllib.parse
from dataclasses import dataclass


class UrlParseError(Exception):
    pass


@dataclass
class ResolvedUrl:
    url_type: str
    graph_endpoint: str
    required_scope: str


_PATTERNS: list[tuple[str, re.Pattern, str, str]] = [
    # Teams channel message: /l/message/19:xxx@thread.tacv2/timestamp
    (
        "channel_message",
        re.compile(
            r"teams\.microsoft\.com/l/message/"
            r"(19:[^/]+@thread\.tacv2)/"
            r"(\d+\.\d+)"
        ),
        "channel_message",
        "ChannelMessage.Read.All",
    ),
    # Teams chat message: /l/message/19:xxx@unq.gbl.spaces/timestamp
    (
        "chat_message",
        re.compile(
            r"teams\.microsoft\.com/l/message/"
            r"(19:[^/]+@(?:unq\.gbl\.spaces|thread\.spaces))/"
            r"(\d+\.\d+)"
        ),
        "chat_message",
        "Chat.Read",
    ),
    # Teams meeting join link: /l/meetup-join/
    (
        "meeting",
        re.compile(r"teams\.microsoft\.com/l/meetup-join/([^/\?&]+)"),
        "meeting",
        "Calendars.Read",
    ),
    # Outlook mail deep link
    (
        "email",
        re.compile(r"outlook\.office(?:365)?\.com/mail/(?:id|deeplink)/([A-Za-z0-9=+/_\-%.]+)"),
        "email",
        "Mail.Read",
    ),
    # SharePoint page (.aspx)
    (
        "sharepoint_page",
        re.compile(r"sharepoint\.com/sites/[^/]+/SitePages/[^/\?]+\.aspx"),
        "sharepoint_page",
        "Sites.Read.All",
    ),
    # OneDrive share link (/_layouts/ pattern) — must come before onedrive_file
    (
        "onedrive_share_link",
        re.compile(r"sharepoint\.com/.*/_layouts/\d+/Doc\.aspx"),
        "onedrive_share_link",
        "Files.Read",
    ),
    # OneDrive personal file (must come after share_link)
    (
        "onedrive_file",
        re.compile(r"sharepoint\.com/personal/[^/]+/(?:Documents|_layouts)"),
        "onedrive_file",
        "Files.Read",
    ),
]


def resolve_url(url: str) -> ResolvedUrl:
    """Parse an M365 URL and return the resolved type + graph endpoint template."""
    for url_type, pattern, endpoint_hint, scope in _PATTERNS:
        m = pattern.search(url)
        if m:
            endpoint = _build_endpoint(url_type, url, m)
            return ResolvedUrl(url_type=url_type, graph_endpoint=endpoint, required_scope=scope)
    raise UrlParseError(f"Unrecognised M365 URL: {url[:100]}")


def _build_endpoint(url_type: str, url: str, match: re.Match) -> str:
    if url_type == "channel_message":
        channel_id = match.group(1)
        message_id = match.group(2).replace(".", "")
        return f"/teams/channels/{channel_id}/messages/{message_id}"
    if url_type == "chat_message":
        chat_id = match.group(1)
        return f"/chats/{chat_id}/messages"
    if url_type == "meeting":
        return "/me/calendarView"
    if url_type == "email":
        msg_id = match.group(1)
        return f"/me/messages/{msg_id}"
    if url_type == "sharepoint_page":
        parsed = urllib.parse.urlparse(url)
        host = parsed.hostname or ""
        path = parsed.path
        return f"/sites/{host}:/{path}"
    if url_type in ("onedrive_file", "onedrive_share_link"):
        return "/me/drive/root"
    return "/"
