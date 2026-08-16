"""resolve composer — parse M365 URLs and fetch their content via Graph."""

import asyncio
import base64
import json
import re
from collections import defaultdict
from datetime import datetime, timedelta, timezone

from ..formatters import (
    _strip_teams_html,
    format_resolved_content_markdown,
    format_section_error,
    graph_dt_to_aware_iso,
)
from ..graph import GraphAPIError, GraphClient
from ..permissions import PermissionRegistry
from ..resolver import ResolvedUrl, UrlParseError, build_chat_thread_url, resolve_url
from ._utils import _error_reason, _escape_odata
from .attachments import body_has_cid, download_attachments, enumerate_attachments


def _encode_share_url(url: str) -> str:
    """Encode a share URL for /shares/{u!<base64>}/driveItem lookup."""
    return "u!" + base64.urlsafe_b64encode(url.encode()).rstrip(b"=").decode()


_RECORDING_UPN_RE = re.compile(r"://([^/]+)/(?::[a-z]:/)?p/([^/]+)/")


def _extract_recording_owner(recording_url: str) -> tuple[str, str]:
    """Pull (host, owner_upn_segment) from a SharePoint recording URL.

    Recording URLs look like:
      https://<tenant>-my.sharepoint.com/:v:/p/<upn-encoded>/<share-token>

    Returns ('', '') if the URL doesn't match.
    """
    if not recording_url:
        return "", ""
    match = _RECORDING_UPN_RE.search(recording_url)
    if not match:
        return "", ""
    return match.group(1), match.group(2)


async def _enrich_call_recording(client: GraphClient, entry: dict) -> None:
    """Add drive_id / drive_item_id / vroom_url / owner_upn to a call entry.

    Mutates the entry in-place. Silent no-op if the entry has no recording_url,
    or if the /shares/ lookup fails (cross-organizer recordings on someone
    else's drive may 403 if the user hasn't opened the recording in Teams UI).
    Graceful degradation is the point — surface what we can, don't fail the
    whole chat-thread render because one drive lookup 403'd.
    """
    recording_url = entry.get("recording_url") or ""
    if not recording_url:
        return

    try:
        share_ref = _encode_share_url(recording_url)
        drive_item = await client.get(f"/shares/{share_ref}/driveItem")
    except GraphAPIError:
        return

    drive_id = (drive_item.get("parentReference") or {}).get("driveId", "")
    drive_item_id = drive_item.get("id", "")
    if not drive_id or not drive_item_id:
        return

    host, owner_upn = _extract_recording_owner(recording_url)
    entry["drive_id"] = drive_id
    entry["drive_item_id"] = drive_item_id
    entry["owner_upn"] = owner_upn
    if host and owner_upn:
        entry["vroom_url"] = (
            f"https://{host}/personal/{owner_upn}"
            f"/_api/v2.0/drives/{drive_id}/items/{drive_item_id}"
        )


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


def _extract_forwarded_message_text(msg: dict) -> str:
    """Extract the inner text of a forwarded-message attachment.

    Teams forwarded messages have body=`<attachment id=...>` and the actual
    content lives in attachments[0].content as a JSON string with key
    `originalMessageContent`. Returns stripped text or '' if not extractable.
    """
    attachments = msg.get("attachments") or []
    for att in attachments:
        if att.get("contentType") != "forwardedMessageReference":
            continue
        raw = att.get("content") or ""
        if not raw:
            continue
        try:
            parsed = json.loads(raw)
        except (json.JSONDecodeError, TypeError):
            continue
        if not isinstance(parsed, dict):
            continue
        inner = parsed.get("originalMessageContent") or ""
        if inner:
            return _strip_teams_html(inner)
    return ""


def _extract_reference_attachments(msg: dict) -> list[dict]:
    """Pull shared files/recordings from a message's reference-type attachments.

    Teams surfaces a shared SharePoint file (including meeting recordings under
    "Microsoft Teams Chat Files") as an attachment with contentType='reference',
    the filename in `name`, and the SharePoint URL in `contentUrl`. Returns a
    list of {name, url} dicts. Attachments missing contentUrl are skipped (not
    renderable). Non-reference types (forwardedMessageReference, inline cards)
    are ignored. Returns [] when none.
    """
    attachments = msg.get("attachments") or []
    result: list[dict] = []
    for att in attachments:
        if att.get("contentType") != "reference":
            continue
        url = att.get("contentUrl") or ""
        if not url:
            continue
        result.append({"name": att.get("name") or url, "url": url})
    return result


def _extract_reply_context(msg: dict) -> dict | None:
    """Pull the quoted-message context from a reply's messageReference attachment.

    A Teams reply/quote carries a `messageReference` attachment whose `content`
    is a JSON string with `messagePreview` (plain-text snippet of the quoted
    message) and `messageSender.user.displayName`. Returns {sender, preview}
    with the preview truncated to 200 chars, or None when there is no reply
    reference or its content is unparseable/empty.
    """
    for att in msg.get("attachments") or []:
        if att.get("contentType") != "messageReference":
            continue
        raw = att.get("content") or ""
        try:
            parsed = json.loads(raw)
        except (json.JSONDecodeError, TypeError):
            return None
        if not isinstance(parsed, dict):
            return None
        preview = (parsed.get("messagePreview") or "").strip()
        if not preview:
            return None
        if len(preview) > 200:
            preview = preview[:200] + "…"
        sender_field = (parsed.get("messageSender") or {}).get("user") or {}
        sender = sender_field.get("displayName") or "Unknown"
        return {"sender": sender, "preview": preview}
    return None


_MARKDOWN_LINK_RE = re.compile(r"\[[^\]]*\]\([^)]*\)")


def _truncate_body(text: str, limit: int) -> str:
    """Truncate `text` to `limit` chars + '…', without cutting a markdown link.

    _strip_teams_html(preserve_links=True) can embed '[label](url)' tokens. A
    naive slice at `limit` could land inside one, leaving a dangling '[label](htt'.
    If the cut point falls inside a link span, back the cut up to the link's start
    so the whole link is dropped rather than mangled. Returns text unchanged when
    already within limit.
    """
    if len(text) <= limit:
        return text
    cut = limit
    for m in _MARKDOWN_LINK_RE.finditer(text):
        if m.start() < cut < m.end():
            cut = m.start()
            break
        if m.start() >= cut:
            break
    return text[:cut].rstrip() + "…"


def _message_entry(msg: dict, name_map: dict[str, str]) -> dict:
    """Classify a real chat message into a typed entry.

    Sender resolution cascades: name_map[user.id] -> user.displayName ->
    application.displayName -> 'Unknown'. Body is stripped via _strip_teams_html
    and truncated to 500 chars + '…'. is_body_empty=True if stripped body is empty.
    Forwarded-message attachments (body=`<attachment id=...>`) have their inner
    text extracted from attachments[0].content as a fallback. Shared files and
    recordings (reference-type attachments) are surfaced in `attachments` as
    {name, url} dicts — even when the body is empty, since the file IS the content.
    Reply/quote context (messageReference attachment) is surfaced in `reply_to`
    as {sender, preview} so a reply is readable on its own.
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
    text = _strip_teams_html(body_content, preserve_links=True)
    if not text:
        text = _extract_forwarded_message_text(msg)
    is_body_empty = not text
    text = _truncate_body(text, 500)

    return {
        "kind": "message",
        "ts": msg.get("createdDateTime", ""),
        "sender": sender,
        "body": text,
        "is_body_empty": is_body_empty,
        "attachments": _extract_reference_attachments(msg),
        "reply_to": _extract_reply_context(msg),
    }


def _event_entry(msg: dict, name_map: dict[str, str] | None = None) -> dict:
    """Classify a non-call system event into a typed entry.

    For member-added/removed events, members[] entries from Graph carry only
    `id` (AAD user object ID) with `displayName: null`. Resolve via name_map
    (built from the chat's expanded members list) before falling back to the
    member's own displayName field.
    """
    detail = msg.get("eventDetail") or {}
    odata_type = detail.get("@odata.type", "")
    ts = msg.get("createdDateTime", "")
    name_map = name_map or {}

    if "membersAdded" in odata_type or "membersDeleted" in odata_type:
        verb = "added" if "membersAdded" in odata_type else "removed"
        event_type = "membersAdded" if "membersAdded" in odata_type else "membersDeleted"
        names: list[str] = []
        for m in detail.get("members") or []:
            user_id = m.get("id") or m.get("userId") or ""
            resolved = name_map.get(user_id) or m.get("displayName")
            if resolved:
                names.append(resolved)
        joined = ", ".join(names) if names else "(someone)"
        return {"kind": "event", "ts": ts, "event_type": event_type, "summary": f"Member {verb}: {joined}"}

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
    entries.extend(_event_entry(e, name_map) for e in other_events)

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
    output_dir: str | None = None,
) -> tuple[dict, str]:
    try:
        resolved = resolve_url(url)
    except UrlParseError as exc:
        markdown = f"❌ Unrecognised URL — {exc}"
        return {"url": url, "kind": "unknown", "data": {}}, markdown

    scope_msg = permissions.check(resolved.required_scope)
    if scope_msg:
        return {"url": url, "kind": resolved.url_type, "data": {}}, scope_msg

    try:
        data = await _fetch_resolved(client, resolved, output_dir)
    except GraphAPIError as exc:
        markdown = format_section_error("Resolve", _error_reason(exc))
        return {"url": url, "kind": resolved.url_type, "data": {}}, markdown

    if "_error" in data:
        markdown = format_section_error("Resolve", data["_error"])
        return {"url": url, "kind": resolved.url_type, "data": {}}, markdown

    structured_data = _build_structured_data(resolved.url_type, data, resolved.extra)
    markdown = format_resolved_content_markdown(
        resolved.url_type, data, chat_url=structured_data.get("chat_url", "")
    )
    return {"url": url, "kind": resolved.url_type, "data": structured_data}, markdown


def _build_structured_data(url_type: str, data: dict, extra: dict | None = None) -> dict:
    """Extract structured fields from the raw Graph response for the given URL type."""
    extra = extra or {}
    if url_type == "email":
        raw_atts = data.get("_attachments") or []
        attachments = [
            {k: v for k, v in a.items() if k not in ("_content_bytes", "kind")}
            for a in raw_atts
        ]
        return {
            "kind": "email",
            "subject": data.get("subject", ""),
            "sender": (data.get("from") or {}).get("emailAddress", {}).get("name", ""),
            "body": (data.get("body") or {}).get("content", "") or (data.get("bodyPreview") or ""),
            "attachments": attachments,
        }
    elif url_type == "chat_thread":
        chat = data.get("chat") or {}
        entries = data.get("entries") or []
        members = chat.get("members") or []
        msg_count = sum(1 for e in entries if e.get("kind") == "message")
        return {
            "kind": "chat_thread",
            "topic": chat.get("topic"),
            "member_count": len(members),
            "recent_message_count": msg_count,
        }
    elif url_type == "chat_message":
        chat_id = extra.get("chat_id", "")
        return {
            "kind": "chat_message",
            "sender": (data.get("from") or {}).get("user", {}).get("displayName", ""),
            "body": (data.get("body") or {}).get("content", ""),
            "created": data.get("createdDateTime"),
            "chat_id": chat_id,
            "chat_url": build_chat_thread_url(chat_id),
        }
    elif url_type == "channel_message":
        return {
            "kind": "channel_message",
            "sender": (data.get("from") or {}).get("user", {}).get("displayName", ""),
            "body": (data.get("body") or {}).get("content", ""),
            "channel_name": None,
        }
    elif url_type == "meeting":
        start_iso, start_tz = graph_dt_to_aware_iso(data.get("start") or {})
        end_iso, end_tz = graph_dt_to_aware_iso(data.get("end") or {})
        return {
            "kind": "meeting",
            "subject": data.get("subject", ""),
            "start": start_iso,
            "end": end_iso,
            "start_timezone": start_tz,
            "end_timezone": end_tz,
        }
    elif url_type == "sharepoint_page":
        return {
            "kind": "sharepoint_page",
            "title": data.get("title") or data.get("displayName") or data.get("name", ""),
            "web_url": data.get("webUrl"),
        }
    elif url_type in ("onedrive_file", "onedrive_share_link"):
        return {
            "kind": "onedrive_file",
            "name": data.get("name", ""),
            "web_url": data.get("webUrl"),
            "size": data.get("size"),
        }
    return {"kind": url_type}


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


async def _fetch_resolved(
    client: GraphClient, resolved: ResolvedUrl, output_dir: str | None = None
) -> dict:
    url_type = resolved.url_type
    endpoint = resolved.graph_endpoint

    if url_type == "email":
        message = await client.get(
            endpoint,
            params={
                "$select": "subject,from,receivedDateTime,bodyPreview,body,toRecipients,ccRecipients,webLink,hasAttachments",
            },
            headers={"Prefer": 'outlook.body-content-type="text"'},
        )
        body_text = (message.get("body") or {}).get("content", "") or (
            message.get("bodyPreview") or ""
        )
        if message.get("hasAttachments") or body_has_cid(body_text):
            entries, enum_error = await enumerate_attachments(client, endpoint)
            if entries and output_dir is not None:
                try:
                    await download_attachments(client, endpoint, entries, output_dir)
                except Exception as exc:
                    message["_attachments_error"] = f"attachment download failed: {exc}"
            message["_attachments"] = entries
            message["_attachments_error"] = message.get("_attachments_error") or enum_error
        else:
            message["_attachments"] = []
            message["_attachments_error"] = None
        return message

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

        # Enrich call entries with drive metadata for recordings. In parallel so
        # multi-recording chats don't serialize the /shares/ lookups. Failures
        # are silent — enrichment is best-effort, chat-thread render must not
        # depend on it.
        recording_entries = [e for e in entries if e.get("kind") == "call" and e.get("recording_url")]
        if recording_entries:
            await asyncio.gather(
                *[_enrich_call_recording(client, e) for e in recording_entries],
                return_exceptions=True,
            )

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
