"""resolve composer — parse M365 URLs and fetch their content via Graph."""

import asyncio
import re
from collections import defaultdict
from datetime import datetime, timedelta, timezone

from ..formatters import _strip_teams_html, format_resolved_content_markdown, format_section_error
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


def _build_member_name_map(chat: dict | None) -> dict[str, str]:
    """Build {userId: displayName} from chat.members. Skip entries lacking either."""
    if not chat:
        return {}
    name_map: dict[str, str] = {}
    for member in chat.get("members") or []:
        user_id = member.get("userId")
        display_name = member.get("displayName")
        if user_id and display_name:
            name_map[user_id] = display_name
    return name_map


def _message_entry(msg: dict, name_map: dict[str, str]) -> dict:
    """Classify a real chat message into a typed entry.

    Sender resolution cascades: name_map[user.id] -> user.displayName ->
    application.displayName -> 'Unknown'. Body is stripped via _strip_teams_html
    and truncated to 500 chars + '…'. is_body_empty=True if stripped body is empty.
    """
    from_field = msg.get("from") or {}
    user_field = from_field.get("user") or {}
    app_field = from_field.get("application") or {}

    user_id = user_field.get("id", "")
    sender = (
        name_map.get(user_id)
        or user_field.get("displayName")
        or app_field.get("displayName")
        or "Unknown"
    )

    body_content = (msg.get("body") or {}).get("content", "")
    text = _strip_teams_html(body_content)
    is_body_empty = not text
    if len(text) > 500:
        text = text[:500] + "…"

    return {
        "kind": "message",
        "ts": msg.get("createdDateTime", ""),
        "sender": sender,
        "body": text,
        "is_body_empty": is_body_empty,
    }


def _event_entry(msg: dict) -> dict:
    """Classify a non-call system event into a typed entry."""
    detail = msg.get("eventDetail") or {}
    odata_type = detail.get("@odata.type", "")
    ts = msg.get("createdDateTime", "")

    if "membersAdded" in odata_type:
        names = [m.get("displayName") for m in (detail.get("members") or []) if m.get("displayName")]
        joined = ", ".join(names) if names else "(someone)"
        return {"kind": "event", "ts": ts, "event_type": "membersAdded", "summary": f"Member added: {joined}"}

    if "membersDeleted" in odata_type:
        names = [m.get("displayName") for m in (detail.get("members") or []) if m.get("displayName")]
        joined = ", ".join(names) if names else "(someone)"
        return {"kind": "event", "ts": ts, "event_type": "membersDeleted", "summary": f"Member removed: {joined}"}

    if "chatRenamed" in odata_type:
        new_name = detail.get("chatDisplayName") or ""
        summary = f'Renamed to "{new_name}"' if new_name else "Chat renamed"
        return {"kind": "event", "ts": ts, "event_type": "chatRenamed", "summary": summary}

    return {"kind": "event", "ts": ts, "event_type": "unknown", "summary": "system event"}


def _group_call_events(call_events: list[dict], name_map: dict[str, str]) -> list[dict]:
    """Group call-related system events by callId into one entry per call.

    Events lacking callId surface as event_type='call_unknown' instead of being
    silently dropped.
    """
    by_call_id: dict[str, list[dict]] = defaultdict(list)
    orphans: list[dict] = []
    for msg in call_events:
        call_id = (msg.get("eventDetail") or {}).get("callId")
        if call_id:
            by_call_id[call_id].append(msg)
        else:
            orphans.append({
                "kind": "event",
                "ts": msg.get("createdDateTime", ""),
                "event_type": "call_unknown",
                "summary": "Call event (no callId)",
            })

    entries: list[dict] = []
    for call_id, events in by_call_id.items():
        events.sort(key=lambda m: m.get("createdDateTime", ""))
        ts_first = events[0].get("createdDateTime", "")
        ts_last = events[-1].get("createdDateTime", "")

        duration: str | None = None
        recording_url = ""
        transcript_ready = False
        initiator: str | None = None
        latest_success_ts = ""

        # Duration resolution: prefer (1) callEnded event's callDuration, then
        # (2) success-status callRecording event, then (3) latest non-zero chunk.
        # PT0S is treated as "no duration yet" (initial event before recording starts).
        recording_durations: list[tuple[str, str, str]] = []  # (ts, status, raw_duration)
        ended_duration: str = ""

        for event in events:
            detail = event.get("eventDetail") or {}
            otype = detail.get("@odata.type", "")
            event_ts = event.get("createdDateTime", "")

            if "callRecording" in otype:
                if not initiator:
                    initiator_field = detail.get("initiator") or {}
                    user_field = initiator_field.get("user") or {}
                    app_field = initiator_field.get("application") or {}
                    user_id = user_field.get("id", "")
                    initiator = (
                        name_map.get(user_id)
                        or user_field.get("displayName")
                        or app_field.get("displayName")
                        or None
                    )
                # recording_url: success-status only, latest wins
                status = (detail.get("callRecordingStatus") or "").lower()
                if status == "success" and detail.get("callRecordingUrl"):
                    if event_ts > latest_success_ts:
                        latest_success_ts = event_ts
                        recording_url = detail["callRecordingUrl"]
                # Track all recording durations for later resolution
                raw_dur = detail.get("callRecordingDuration") or ""
                if raw_dur and raw_dur != "PT0S":
                    recording_durations.append((event_ts, status, raw_dur))

            elif "callTranscript" in otype:
                transcript_ready = True

            elif "callEnded" in otype:
                if detail.get("callDuration"):
                    ended_duration = detail["callDuration"]

        # Apply duration precedence after iterating all events.
        if ended_duration:
            duration = _parse_iso_duration(ended_duration) or None
        elif recording_durations:
            success_durations = [(ts, raw) for ts, status, raw in recording_durations if status == "success"]
            chosen = max(success_durations, key=lambda x: x[0])[1] if success_durations else max(recording_durations, key=lambda x: x[0])[2]
            duration = _parse_iso_duration(chosen) or None

        entries.append({
            "kind": "call",
            "ts": ts_first,
            "end_ts": ts_last,
            "duration": duration,
            "recording_url": recording_url,
            "transcript_ready": transcript_ready,
            "initiator": initiator,
        })

    return entries + orphans


_CALL_RELATED_FRAGMENTS = ("callRecording", "callTranscript", "callStarted", "callEnded")


def _is_call_related_event(msg: dict) -> bool:
    detail = msg.get("eventDetail") or {}
    odata_type = detail.get("@odata.type", "")
    return any(frag in odata_type for frag in _CALL_RELATED_FRAGMENTS)


def _normalize_chat_entries(messages: list[dict], name_map: dict[str, str]) -> list[dict]:
    """Classify, group, sort, and cap raw chat messages into typed entries."""
    real_messages: list[dict] = []
    call_events: list[dict] = []
    other_events: list[dict] = []

    for msg in messages:
        if msg.get("eventDetail"):
            if _is_call_related_event(msg):
                call_events.append(msg)
            else:
                other_events.append(msg)
        else:
            real_messages.append(msg)

    entries: list[dict] = []
    entries.extend(_message_entry(m, name_map) for m in real_messages)
    entries.extend(_group_call_events(call_events, name_map))
    entries.extend(_event_entry(e) for e in other_events)

    entries.sort(key=lambda e: e.get("ts", ""), reverse=True)
    return entries[:25]


async def _paginate_chat_messages(
    client,
    chat_id: str,
    max_messages: int = 100,
    max_pages: int = 3,
) -> tuple[list[dict], str | None]:
    """Fetch up to max_messages chat messages, following @odata.nextLink.

    Returns (messages, partial_error). partial_error is None on full success
    or a reason string if a non-first-page request failed (in which case the
    list contains pages successfully fetched before the failure).

    Raises GraphAPIError on first-page failure (caller surfaces as full error).
    """
    messages: list[dict] = []
    seen_ids: set[str] = set()
    next_url: str | None = f"/chats/{chat_id}/messages?$top=50"
    pages = 0
    partial_error: str | None = None

    while next_url and pages < max_pages and len(messages) < max_messages:
        try:
            response = await client.get(next_url)
        except GraphAPIError as exc:
            if pages == 0:
                raise
            partial_error = f"page {pages + 1} failed: {_error_reason(exc)}"
            break
        for msg in response.get("value", []):
            msg_id = msg.get("id")
            if msg_id is None:
                # Defensive: include unidentified messages without dedup.
                messages.append(msg)
            elif msg_id not in seen_ids:
                seen_ids.add(msg_id)
                messages.append(msg)
        next_url = response.get("@odata.nextLink") or None
        pages += 1

    return messages[:max_messages], partial_error


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
        msgs_task = _paginate_chat_messages(client, chat_id)
        chat_result, msgs_result = await asyncio.gather(
            chat_task, msgs_task, return_exceptions=True
        )

        chat = chat_result if not isinstance(chat_result, BaseException) else None

        if isinstance(msgs_result, BaseException):
            messages: list[dict] = []
            msgs_error_reason: str | None = _error_reason(msgs_result)
        else:
            messages, partial_error = msgs_result
            msgs_error_reason = partial_error

        name_map = _build_member_name_map(chat) if chat else {}
        entries = _normalize_chat_entries(messages, name_map)

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
            "entries": entries,
            "meeting": meeting_event,
            "_chat_error": _error_reason(chat_result) if isinstance(chat_result, BaseException) else None,
            "_messages_error": msgs_error_reason,
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
