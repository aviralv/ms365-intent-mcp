"""meeting composer — full context for a single meeting."""

import re
import urllib.parse
from datetime import datetime, timedelta

from ..formatters import format_event_detail_markdown, graph_dt_to_aware_iso
from ..graph import GraphAPIError, GraphClient
from ..permissions import PermissionRegistry
from ._utils import _escape_odata
from .resolve import _enrich_call_recording

_JOIN_THREAD_RE = re.compile(r"meetup-join/(19:[^/]+@thread\.v2)")


async def compose_meeting(
    client: GraphClient,
    permissions: PermissionRegistry,
    identifier: str,
    timezone: str,
) -> tuple[dict, str]:
    event = await _resolve_event(client, identifier, timezone)
    if event is None:
        markdown = f'❌ No meeting found matching "{identifier}"'
        return {"id": "", "subject": identifier, "start": "", "end": "", "organizer": {}, "attendees": [], "recording": None}, markdown

    recording = await _resolve_recording_for_event(client, event)
    if recording:
        event = {**event, "_recording": recording}

    markdown = format_event_detail_markdown(event)

    # Build structured data dict from the event
    organizer_email_addr = event.get("organizer", {}).get("emailAddress", {})
    organizer_data = {
        "name": organizer_email_addr.get("name", "Unknown"),
        "email": organizer_email_addr.get("address") or None,
    }

    attendees_data = []
    for a in event.get("attendees", []):
        ea = a.get("emailAddress", {})
        attendees_data.append({
            "name": ea.get("name", ""),
            "email": ea.get("address") or None,
            "response": a.get("status", {}).get("response", "none"),
        })

    online_meeting_data = None
    if event.get("isOnlineMeeting"):
        join_url = (event.get("onlineMeeting") or {}).get("joinUrl", "")
        if join_url:
            online_meeting_data = {"join_url": join_url}

    recording_data = None
    if recording:
        recording_data = {
            "recording_url": recording.get("recording_url", ""),
            "display_name": recording.get("display_name"),
            "transcript_ready": recording.get("transcript_ready", False),
            "drive_id": recording.get("drive_id"),
            "drive_item_id": recording.get("drive_item_id"),
            "owner_upn": recording.get("owner_upn"),
            "vroom_url": recording.get("vroom_url"),
        }

    start_iso, start_tz = graph_dt_to_aware_iso(event.get("start", {}))
    end_iso, end_tz = graph_dt_to_aware_iso(event.get("end", {}))
    data = {
        "id": event.get("id", ""),
        "subject": event.get("subject", ""),
        "start": start_iso,
        "end": end_iso,
        "start_timezone": start_tz,
        "end_timezone": end_tz,
        "organizer": organizer_data,
        "attendees": attendees_data,
        "location": event.get("location", {}).get("displayName") or None,
        "online_meeting": online_meeting_data,
        "recording": recording_data,
    }
    return data, markdown


async def _resolve_recording_for_event(client: GraphClient, event: dict) -> dict | None:
    """Look up recording metadata for an event via its chat thread.

    Graph doesn't populate `onlineMeeting.chatInfo` on /me/calendarView and
    /me/onlineMeetings requires a scope we don't have — so we extract the
    thread id from the joinUrl (verified live: every Teams `meetup-join`
    URL carries `19:<thread>@thread.v2` in its path).

    Returns None when: not online, no joinUrl, no thread parseable, chat
    fetch fails, no callRecording events, or no success-status recording.
    All are 'no recording surfaced' from the caller's perspective.
    """
    if not event.get("isOnlineMeeting"):
        return None
    join_url = ((event.get("onlineMeeting") or {}).get("joinUrl") or "")
    if not join_url:
        return None
    join_url_decoded = urllib.parse.unquote(join_url)
    match = _JOIN_THREAD_RE.search(join_url_decoded)
    if not match:
        return None
    thread_id = match.group(1)

    try:
        msgs = await client.get(f"/chats/{thread_id}/messages", params={"$top": "50"})
    except GraphAPIError:
        return None

    messages = (msgs or {}).get("value", [])
    entry = _extract_recording_entry(messages)
    if not entry:
        return None

    # Enrich in-place with drive metadata via /shares/. Best-effort — a 403
    # (cross-organizer, not-yet-opened recording) still leaves the URL surfaced.
    await _enrich_call_recording(client, entry)
    return entry


def _extract_recording_entry(messages: list[dict]) -> dict | None:
    """Collapse callRecording events into a single entry with the freshest
    success URL, transcript_ready flag, and initiator. Returns None if no
    callRecording events are present."""
    recording_events = [
        m for m in messages
        if "callRecording" in ((m.get("eventDetail") or {}).get("@odata.type", ""))
    ]
    transcript_events = [
        m for m in messages
        if "callTranscript" in ((m.get("eventDetail") or {}).get("@odata.type", ""))
    ]
    if not recording_events:
        return None

    recording_url = ""
    latest_success_ts = ""
    display_name = ""
    for msg in recording_events:
        detail = msg.get("eventDetail") or {}
        status = (detail.get("callRecordingStatus") or "").lower()
        event_ts = msg.get("createdDateTime", "")
        if status == "success" and detail.get("callRecordingUrl"):
            if event_ts > latest_success_ts:
                latest_success_ts = event_ts
                recording_url = detail["callRecordingUrl"]
                display_name = detail.get("callRecordingDisplayName") or display_name

    if not recording_url:
        return None

    return {
        "recording_url": recording_url,
        "display_name": display_name,
        "transcript_ready": bool(transcript_events),
    }


async def _resolve_event(
    client: GraphClient,
    identifier: str,
    timezone: str,
) -> dict | None:
    if identifier.lower() == "next":
        return await _find_next_event(client, timezone)

    if " " not in identifier:
        try:
            headers = GraphClient.calendar_headers(timezone)
            return await client.get(f"/me/events/{identifier}", headers=headers)
        except GraphAPIError:
            pass

    return await _search_by_subject(client, identifier, timezone)


async def _find_next_event(client: GraphClient, timezone: str) -> dict | None:
    now = datetime.now()
    start_iso = now.strftime("%Y-%m-%dT%H:%M:%S")
    end_iso = (now + timedelta(days=1)).strftime("%Y-%m-%dT23:59:59")

    headers = GraphClient.calendar_headers(timezone)
    result = await client.get("/me/calendarView", params={
        "startDateTime": start_iso,
        "endDateTime": end_iso,
        "$orderby": "start/dateTime",
        "$top": "1",
    }, headers=headers)

    events = result.get("value", [])
    return events[0] if events else None


async def _search_by_subject(
    client: GraphClient,
    subject: str,
    timezone: str,
) -> dict | None:
    now = datetime.now()
    start_iso = (now - timedelta(days=7)).strftime("%Y-%m-%dT00:00:00")
    end_iso = (now + timedelta(days=7)).strftime("%Y-%m-%dT23:59:59")

    headers = GraphClient.calendar_headers(timezone)

    try:
        result = await client.get("/me/calendarView", params={
            "startDateTime": start_iso,
            "endDateTime": end_iso,
            "$orderby": "start/dateTime asc",
            "$top": "50",
            "$filter": f"contains(subject, '{_escape_odata(subject)}')",
        }, headers=headers)

        events = result.get("value", [])
        if events:
            return events[0]
    except GraphAPIError:
        pass

    result = await client.get("/me/calendarView", params={
        "startDateTime": start_iso,
        "endDateTime": end_iso,
        "$orderby": "start/dateTime asc",
        "$top": "50",
    }, headers=headers)

    subject_lower = subject.lower()
    for event in result.get("value", []):
        if subject_lower in event.get("subject", "").lower():
            return event

    return None
