"""resolve composer — parse M365 URLs and fetch their content via Graph."""

import asyncio
import re
from datetime import datetime, timedelta, timezone

from ..formatters import format_resolved_content_markdown, format_section_error
from ..graph import GraphAPIError, GraphClient
from ..permissions import PermissionRegistry
from ..resolver import ResolvedUrl, UrlParseError, resolve_url
from ._utils import _error_reason, _escape_odata


_ISO_DURATION_RE = re.compile(
    r"^PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)(?:\.\d+)?S)?$"
)


def _parse_iso_duration(iso: str) -> str:
    """Convert ISO-8601 duration like 'PT25M38.4S' to compact 'h/m/s' form.

    Returns empty string if input is empty or doesn't match.
    Handles only PT-form durations (Graph call durations always are);
    day component (P1DT...) and other prefixes return empty.
    """
    if not iso:
        return ""
    m = _ISO_DURATION_RE.match(iso)
    if not m:
        return ""
    hours, minutes, seconds = m.groups()
    parts = []
    if hours:
        parts.append(f"{int(hours)}h")
    if minutes:
        parts.append(f"{int(minutes)}m")
    if seconds:
        parts.append(f"{int(seconds)}s")
    return "".join(parts) if parts else ""


async def compose_resolve(
    client: GraphClient,
    permissions: PermissionRegistry,
    url: str,
) -> str:
    try:
        resolved = resolve_url(url)
    except UrlParseError as exc:
        return f"❌ Unrecognised URL — {exc}"

    scope_msg = permissions.check(resolved.required_scope)
    if scope_msg:
        return scope_msg

    try:
        data = await _fetch_resolved(client, resolved)
    except GraphAPIError as exc:
        return format_section_error("Resolve", _error_reason(exc))

    if "_error" in data:
        return format_section_error("Resolve", data["_error"])

    return format_resolved_content_markdown(resolved.url_type, data)


async def _find_meeting_by_join_url(client: GraphClient, join_url: str) -> dict | None:
    """Search the user's calendar (±14 days) for an event whose joinUrl contains the given fragment.

    Returns the matched event dict, or None if no match.
    """
    if not join_url:
        return None
    now = datetime.now(timezone.utc)
    start = (now - timedelta(days=14)).strftime("%Y-%m-%dT%H:%M:%SZ")
    end = (now + timedelta(days=14)).strftime("%Y-%m-%dT%H:%M:%SZ")
    try:
        result = await client.get("/me/calendarView", params={
            "startDateTime": start,
            "endDateTime": end,
            "$top": "50",
            "$select": "subject,start,end,organizer,attendees,body,location,isOnlineMeeting,onlineMeeting",
        })
    except GraphAPIError:
        return None
    events = (result or {}).get("value", [])
    for event in events:
        event_join = (event.get("onlineMeeting") or {}).get("joinUrl", "")
        if join_url in event_join:
            return event
    return None


async def _get_event_by_id(client: GraphClient, event_id: str) -> dict | None:
    """Direct calendar event lookup by ID. Returns None on empty ID or any GraphAPIError.

    The calendar fuse is opportunistic — a 404 (event deleted, or on someone else's calendar)
    is silently swallowed so the chat_thread response still renders.
    """
    if not event_id:
        return None
    try:
        return await client.get(f"/me/events/{event_id}", params={
            "$select": "subject,start,end,organizer,attendees,body,location,isOnlineMeeting,onlineMeeting",
        })
    except GraphAPIError:
        return None


async def _fetch_resolved(client: GraphClient, resolved: ResolvedUrl) -> dict:
    url_type = resolved.url_type
    endpoint = resolved.graph_endpoint

    if url_type == "email":
        return await client.get(endpoint, params={
            "$select": "subject,from,receivedDateTime,bodyPreview,body",
        })

    elif url_type == "channel_message":
        return await client.get(endpoint, params={
            "$select": "body,from,createdDateTime,subject",
        })

    elif url_type == "chat_message":
        return await client.get(endpoint, params={
            "$select": "body,from,createdDateTime",
        })

    elif url_type == "chat_thread":
        chat_id = resolved.extra["chat_id"]
        chat_task = client.get(f"/chats/{chat_id}", params={
            "$select": "id,topic,chatType,webUrl,onlineMeetingInfo",
            "$expand": "members",
        })
        msgs_task = client.get(f"/chats/{chat_id}/messages", params={
            "$top": "20",
        })
        chat_result, msgs_result = await asyncio.gather(
            chat_task, msgs_task, return_exceptions=True
        )

        chat = chat_result if not isinstance(chat_result, BaseException) else None
        messages_raw = (
            msgs_result.get("value", [])
            if isinstance(msgs_result, dict) else []
        )
        messages = sorted(
            messages_raw,
            key=lambda m: m.get("createdDateTime", ""),
            reverse=True,
        )[:20]

        meeting_event = None
        if chat:
            omi = chat.get("onlineMeetingInfo") or {}
            calendar_event_id = omi.get("calendarEventId") or ""
            join_web_url = omi.get("joinWebUrl") or ""
            if calendar_event_id:
                meeting_event = await _get_event_by_id(client, calendar_event_id)
            elif join_web_url:
                meeting_event = await _find_meeting_by_join_url(client, join_web_url)

        return {
            "chat": chat,
            "messages": messages,
            "meeting": meeting_event,
            "_chat_error": _error_reason(chat_result) if isinstance(chat_result, BaseException) else None,
            "_messages_error": _error_reason(msgs_result) if isinstance(msgs_result, BaseException) else None,
            "_url_type": "chat_thread",
        }

    elif url_type == "meeting":
        thread_id = resolved.extra.get("thread_id", "")
        event = await _find_meeting_by_join_url(client, thread_id)
        if event is None:
            return {"_error": "No matching meeting found for this Teams link."}
        return event

    elif url_type in ("onedrive_file", "onedrive_share_link"):
        return await client.get(endpoint, params={
            "$select": "name,size,webUrl,lastModifiedDateTime,createdDateTime,file",
        })

    elif url_type == "sharepoint_page":
        site_data = await client.get(endpoint)
        site_id = (site_data or {}).get("id", "")
        page_filename = resolved.extra.get("page_filename", "")
        if site_id and page_filename:
            try:
                page_data = await _fetch_sharepoint_page(client, site_id, page_filename)
                if page_data:
                    page_data["_page_found"] = True
                    page_data["_site_name"] = (site_data or {}).get("displayName", "")
                    return page_data
            except GraphAPIError:
                pass
        return site_data

    else:
        return {}


async def _fetch_sharepoint_page(client: GraphClient, site_id: str, filename: str) -> dict | None:
    """Look up a SharePoint page by filename via the Site Pages list."""
    lists_result = await client.get(
        f"/sites/{site_id}/lists",
        params={"$filter": "displayName eq 'Site Pages'", "$select": "id"},
    )
    lists = (lists_result or {}).get("value", [])
    if not lists:
        return None
    list_id = lists[0]["id"]
    items_result = await client.get(
        f"/sites/{site_id}/lists/{list_id}/items",
        params={
            "$filter": f"fields/FileLeafRef eq '{_escape_odata(filename)}'",
            "$select": "id,webUrl",
            "$expand": "fields($select=FileLeafRef,Title,Modified)",
            "$top": "1",
        },
        headers={"Prefer": "HonorNonIndexedQueriesWarningMayFailRandomly"},
    )
    items = (items_result or {}).get("value", [])
    if not items:
        return None
    item = items[0]
    fields = item.get("fields", {})
    return {
        "name": fields.get("FileLeafRef", filename),
        "title": fields.get("Title", ""),
        "webUrl": item.get("webUrl", ""),
        "lastModifiedDateTime": fields.get("Modified", ""),
    }
