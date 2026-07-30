"""URL resolver — regex dispatch table for 7 M365 URL types."""

from __future__ import annotations

import base64
import json
import re
import urllib.parse
from dataclasses import dataclass, field


class UrlParseError(Exception):
    pass


def normalize_message_id(msg_id: str) -> str:
    """Normalize an Outlook message/event ID to the form Graph's REST paths accept.

    Outlook exposes IDs in two encodings. Client "copy link" (mail/id/<id>) and
    Search `hitId` use a URL-safe form ('-'/'_'). OWA (owa/?ItemID=<id>) uses
    standard base64 ('+'/'/'). Graph's /me/messages/{id} and /me/events/{id}
    accept the URL-safe form only — a raw '/' spliced into the path (even
    percent-encoded to '%2F') 404s with RequestBroker ParseUri. Map '/'→'-' and
    '+'→'_'. Idempotent on already-URL-safe IDs (no '/' or '+' to replace), so
    it is safe to apply at every ID-consuming boundary.
    """
    return msg_id.replace("/", "-").replace("+", "_")


@dataclass
class ResolvedUrl:
    url_type: str
    graph_endpoint: str
    required_scope: str
    extra: dict[str, str] = field(default_factory=dict)


# Base of the Teams "deep link to a chat" URL format. Single source of truth
# for both the chat_thread parse pattern below and build_chat_thread_url() —
# keeping parse and build next to each other so the l/chat/ format can't drift.
_CHAT_THREAD_URL_BASE = "https://teams.microsoft.com/l/chat/"


def build_chat_thread_url(chat_id: str) -> str:
    """Construct a Teams chat-thread deep link from a chat_id.

    Inverse of the chat_thread parse pattern. The reconstructed URL round-trips
    through resolve_url() but lacks the ?tenantId= query param that Graph's real
    webUrl carries — that only matters for cross-tenant/guest scenarios and would
    require an extra /chats/{id} fetch to obtain. Returns '' for a falsy chat_id.
    """
    return f"{_CHAT_THREAD_URL_BASE}{chat_id}" if chat_id else ""


_PATTERNS: list[tuple[str, re.Pattern, str]] = [
    (
        "channel_message",
        re.compile(
            r"teams\.microsoft\.com/l/message/"
            r"(19:[^/]+@thread\.tacv2)/"
            r"(\d+(?:\.\d+)?)"
        ),
        "ChannelMessage.Read.All",
    ),
    (
        "chat_message",
        re.compile(
            r"teams\.microsoft\.com/l/message/"
            r"(19:[^/]+@(?:unq\.gbl\.spaces|thread\.spaces|thread\.v2))/"
            r"(\d+(?:\.\d+)?)"
        ),
        "Chat.ReadWrite",
    ),
    (
        "chat_thread",
        re.compile(
            r"teams\.microsoft\.com/l/chat/"
            r"(19:[^/]+@(?:thread\.v2|unq\.gbl\.spaces|thread\.spaces))"
        ),
        "Chat.ReadWrite",
    ),
    (
        "meeting",
        re.compile(r"teams\.microsoft\.com/l/meetup-join/([^/\?&]+)"),
        "Calendars.Read",
    ),
    (
        "email",
        re.compile(r"outlook\.office(?:365)?\.com/mail/(?:id|deeplink)/([A-Za-z0-9=+/_\-%.]+)"),
        "Mail.Read",
    ),
    (
        "email",
        re.compile(r"outlook\.office(?:365)?\.com/owa/\?[^\s]*?ItemID=([A-Za-z0-9=+/_\-%.]+)"),
        "Mail.Read",
    ),
    (
        "sharepoint_page",
        re.compile(r"sharepoint\.com/sites/[^/]+/SitePages/[^/\?]+\.aspx"),
        "Sites.Read.All",
    ),
    (
        "onedrive_share_link",
        re.compile(r"sharepoint\.com/.*/_layouts/\d+/Doc\.aspx"),
        "Files.Read",
    ),
    (
        "onedrive_file",
        re.compile(r"sharepoint\.com/personal/[^/]+/(?:Documents|_layouts)"),
        "Files.Read",
    ),
]


def resolve_url(url: str) -> ResolvedUrl:
    """Parse an M365 URL and return the resolved type + graph endpoint template."""
    url = urllib.parse.unquote(url)
    for url_type, pattern, scope in _PATTERNS:
        m = pattern.search(url)
        if m:
            extra = _build_extra(url_type, url, m)
            endpoint = _build_endpoint(url_type, url, m, extra)
            return ResolvedUrl(
                url_type=url_type,
                graph_endpoint=endpoint,
                required_scope=scope,
                extra=extra,
            )
    raise UrlParseError(f"Unrecognised M365 URL: {url[:100]}")


def _encode_sharing_url(url: str) -> str:
    """Encode a sharing URL for the /shares/ API (base64url with u! prefix)."""
    encoded = base64.urlsafe_b64encode(url.encode()).rstrip(b"=").decode()
    return f"u!{encoded}"


def _parse_onedrive_personal_path(url: str) -> tuple[str, str]:
    """Extract UPN and relative file path from a OneDrive personal URL.

    URL format: https://{tenant}-my.sharepoint.com/personal/{upn_encoded}/Documents/{path}
    UPN encoding: dots→underscores, @ → underscore, so user_company_com → user@company.com
    """
    parsed = urllib.parse.urlparse(url)
    path = parsed.path
    match = re.search(r"/personal/([^/]+)/Documents/(.+)", path)
    if not match:
        match = re.search(r"/personal/([^/]+)/", path)
        if match:
            return _decode_upn(match.group(1)), ""
        return "", ""
    upn_encoded = match.group(1)
    relative_path = urllib.parse.unquote(match.group(2))
    return _decode_upn(upn_encoded), relative_path


_MULTI_PART_TLDS = {"co.uk", "com.au", "co.jp", "co.nz", "com.br", "co.in", "org.uk", "ac.uk"}


def _decode_upn(encoded: str) -> str:
    """Decode a OneDrive personal UPN from URL path segment.

    SharePoint encodes user@domain as: dots→underscores, @→underscore.
    e.g. alice_smith_example_com → alice.smith@example.com
         alice_smith_example_co_uk → alice.smith@example.co.uk
    """
    parts = encoded.split("_")
    if len(parts) < 3:
        return encoded
    if len(parts) >= 4:
        potential_tld = f"{parts[-2]}.{parts[-1]}"
        if potential_tld in _MULTI_PART_TLDS:
            domain = ".".join(parts[-3:])
            username = ".".join(parts[:-3])
            if username:
                return f"{username}@{domain}"
    domain = ".".join(parts[-2:])
    username = ".".join(parts[:-2])
    return f"{username}@{domain}"


def _build_extra(url_type: str, url: str, match: re.Match) -> dict[str, str]:
    """Extract additional context needed by the composer."""
    if url_type == "chat_thread":
        return {"chat_id": match.group(1)}
    if url_type == "chat_message":
        return {"chat_id": match.group(1)}
    if url_type == "meeting":
        return {"thread_id": urllib.parse.unquote(match.group(1))}
    if url_type == "sharepoint_page":
        parsed = urllib.parse.urlparse(url)
        page_match = re.search(r"/SitePages/([^/?]+\.aspx)", parsed.path)
        page_filename = page_match.group(1) if page_match else ""
        return {"page_filename": page_filename}
    if url_type == "channel_message":
        extra: dict[str, str] = {
            "channel_id": match.group(1),
            "message_id": match.group(2),
        }
        parsed = urllib.parse.urlparse(url)
        qs = urllib.parse.parse_qs(parsed.query)
        context_raw = qs.get("context", [""])[0]
        if context_raw:
            try:
                ctx = json.loads(context_raw)
                group_id = ctx.get("groupId", "")
                if group_id:
                    extra["group_id"] = group_id
            except (json.JSONDecodeError, TypeError):
                pass
        return extra
    return {}


def _build_endpoint(url_type: str, url: str, match: re.Match, extra: dict[str, str]) -> str:
    if url_type == "chat_thread":
        return f"/chats/{match.group(1)}"
    if url_type == "channel_message":
        channel_id = extra.get("channel_id", match.group(1))
        message_id = extra.get("message_id", match.group(2))
        group_id = extra.get("group_id", "")
        if group_id:
            return f"/teams/{group_id}/channels/{channel_id}/messages/{message_id}"
        return f"/chats/{channel_id}/messages/{message_id}"
    if url_type == "chat_message":
        chat_id = match.group(1)
        message_id = match.group(2)
        return f"/chats/{chat_id}/messages/{message_id}"
    if url_type == "meeting":
        return "/me/calendarView"
    if url_type == "email":
        # `resolve_url` already unquoted the incoming URL, so any '%2F' is now
        # a raw '/' at this point.
        return f"/me/messages/{urllib.parse.quote(normalize_message_id(match.group(1)), safe='')}"
    if url_type == "sharepoint_page":
        parsed = urllib.parse.urlparse(url)
        host = parsed.hostname or ""
        path_parts = parsed.path.split("/SitePages/")
        site_path = path_parts[0] if path_parts else parsed.path
        return f"/sites/{host}:{site_path}"
    if url_type == "onedrive_share_link":
        return f"/shares/{_encode_sharing_url(url)}/driveItem"
    if url_type == "onedrive_file":
        upn, relative_path = _parse_onedrive_personal_path(url)
        if upn and relative_path:
            return f"/users/{upn}/drive/root:/{relative_path}"
        if upn:
            return f"/users/{upn}/drive/root"
        return "/me/drive/root"
    return "/"
