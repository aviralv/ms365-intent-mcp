"""URL resolver — regex dispatch table for 7 M365 URL types."""

from __future__ import annotations

import base64
import json
import re
import urllib.parse
from dataclasses import dataclass, field


class UrlParseError(Exception):
    pass


@dataclass
class ResolvedUrl:
    url_type: str
    graph_endpoint: str
    required_scope: str
    extra: dict[str, str] = field(default_factory=dict)


_PATTERNS: list[tuple[str, re.Pattern, str]] = [
    (
        "channel_message",
        re.compile(
            r"teams\.microsoft\.com/l/message/"
            r"(19:[^/]+@thread\.tacv2)/"
            r"(\d+\.\d+)"
        ),
        "ChannelMessage.Read.All",
    ),
    (
        "chat_message",
        re.compile(
            r"teams\.microsoft\.com/l/message/"
            r"(19:[^/]+@(?:unq\.gbl\.spaces|thread\.spaces))/"
            r"(\d+\.\d+)"
        ),
        "Chat.Read",
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
    for url_type, pattern, scope in _PATTERNS:
        m = pattern.search(url)
        if m:
            endpoint = _build_endpoint(url_type, url, m)
            extra = _build_extra(url_type, url, m)
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


def _decode_upn(encoded: str) -> str:
    """Decode a OneDrive personal UPN from URL path segment.

    Convention: last two underscores represent @ and domain separator.
    e.g. aviral_vaid_sap_com → aviral.vaid@sap.com
    """
    parts = encoded.split("_")
    if len(parts) >= 3:
        domain = ".".join(parts[-2:])
        username = ".".join(parts[:-2])
        return f"{username}@{domain}"
    return encoded


def _build_extra(url_type: str, url: str, match: re.Match) -> dict[str, str]:
    """Extract additional context needed by the composer."""
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


def _build_endpoint(url_type: str, url: str, match: re.Match) -> str:
    if url_type == "channel_message":
        channel_id = match.group(1)
        message_id = match.group(2)
        parsed = urllib.parse.urlparse(url)
        qs = urllib.parse.parse_qs(parsed.query)
        context_raw = qs.get("context", [""])[0]
        group_id = ""
        if context_raw:
            try:
                ctx = json.loads(context_raw)
                group_id = ctx.get("groupId", "")
            except (json.JSONDecodeError, TypeError):
                pass
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
        msg_id = match.group(1)
        return f"/me/messages/{msg_id}"
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
