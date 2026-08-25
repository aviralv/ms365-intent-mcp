"""meeting composer — full context for a single meeting."""

import re
import urllib.parse
from datetime import datetime, timedelta

from ..formatters import format_event_detail_markdown, graph_dt_to_aware_iso
from ..graph import GraphAPIError, GraphClient
from ..permissions import PermissionRegistry
from ..transcripts import TEAMS_FILENAME_RE
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
        return {
            "id": "",
            "subject": identifier,
            "start": "",
            "end": "",
            "organizer": {},
            "attendees": [],
            "recording": None,
        }, markdown

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
        attendees_data.append(
            {
                "name": ea.get("name", ""),
                "email": ea.get("address") or None,
                "response": a.get("status", {}).get("response", "none"),
            }
        )

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
            "recording_date": recording.get("recording_date"),
            "occurrence_date": recording.get("occurrence_date"),
            "date_matches_occurrence": recording.get("date_matches_occurrence"),
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
    join_url = (event.get("onlineMeeting") or {}).get("joinUrl") or ""
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
    occurrence_date = _event_occurrence_date(event)
    entry = _extract_recording_entry(messages, occurrence_date)
    if not entry:
        return None

    # Enrich in-place with drive metadata via /shares/. Best-effort — a 403
    # (cross-organizer, not-yet-opened recording) still leaves the URL surfaced.
    await _enrich_call_recording(client, entry)
    return entry


def _event_occurrence_date(event: dict) -> str:
    """Local (timezone-aware) calendar date of the event occurrence, ``YYYY-MM-DD``.

    Returns "" when the event has no resolvable start — callers then skip
    date-matching and fall back to freshest-recording selection.

    Note: this is the event's *own* timezone date (Graph populates
    ``start.timeZone`` under the outlook.timezone Prefer header), compared
    downstream against the Teams filename's *tenant-local* date. When the event
    timezone and tenant timezone differ enough to straddle midnight, a correct
    recording can be flagged ``date_matches_occurrence=False`` (a false-stale
    warning). That's the safe failure mode — visible over-warning, never silent
    wrong data — and rare for single-tenant same-tz usage.
    """
    start_iso, _ = graph_dt_to_aware_iso(event.get("start", {}))
    return start_iso[:10] if start_iso else ""


def _recording_date_from_name(display_name: str, fallback_ts: str) -> str:
    """Date the recording was actually made, ``YYYY-MM-DD``.

    Parsed from the Teams-generated filename (``-YYYYMMDD_HHMMSS-Meeting ...``),
    which is the true meeting date. Falls back to the chat event's
    ``createdDateTime`` when the display name doesn't follow the convention.
    """
    match = TEAMS_FILENAME_RE.match(display_name or "")
    if match:
        d = match.group(2)
        return f"{d[:4]}-{d[4:6]}-{d[6:8]}"
    return (fallback_ts or "")[:10]


def _extract_recording_entry(
    messages: list[dict], occurrence_date: str = ""
) -> dict | None:
    """Collapse callRecording events into a single entry for this occurrence.

    Recurring meetings share one chat thread across all occurrences, so the
    thread carries callRecording events for *every* recorded occurrence. We
    can't pick 'freshest overall' — that returns a prior occurrence's file
    (issue #79). Instead we match on the date embedded in each recording's
    Teams filename:

      * If a recording's date matches ``occurrence_date``, choose the freshest
        of those (``date_matches_occurrence=True``).
      * Otherwise fall back to the freshest recording in the thread and flag
        ``date_matches_occurrence=False`` with its own ``recording_date`` so a
        stale pick is a hard, visible signal, not silent wrong data (issue #51).

    ``date_matches_occurrence`` is ``None`` when ``occurrence_date`` is unknown
    (the match can't be evaluated). Returns None if no success recording exists.
    """
    recording_events = [
        m
        for m in messages
        if "callRecording" in ((m.get("eventDetail") or {}).get("@odata.type", ""))
    ]
    transcript_events = [
        m
        for m in messages
        if "callTranscript" in ((m.get("eventDetail") or {}).get("@odata.type", ""))
    ]
    if not recording_events:
        return None

    # (event_ts, recording_url, display_name, recording_date, call_id) per success event.
    candidates: list[tuple[str, str, str, str, str]] = []
    for msg in recording_events:
        detail = msg.get("eventDetail") or {}
        status = (detail.get("callRecordingStatus") or "").lower()
        url = detail.get("callRecordingUrl")
        if status != "success" or not url:
            continue
        event_ts = msg.get("createdDateTime", "")
        display_name = detail.get("callRecordingDisplayName") or ""
        rec_date = _recording_date_from_name(display_name, event_ts)
        call_id = detail.get("callId") or ""
        candidates.append((event_ts, url, display_name, rec_date, call_id))

    if not candidates:
        return None

    same_date = [c for c in candidates if occurrence_date and c[3] == occurrence_date]
    pool = same_date or candidates
    # Freshest by chat-event timestamp within the chosen pool.
    _, recording_url, display_name, recording_date, call_id = max(pool, key=lambda c: c[0])

    date_matches: bool | None
    if not occurrence_date:
        date_matches = None
    else:
        date_matches = bool(same_date)

    # Tie transcript_ready to the *chosen* recording via callId — a shared
    # recurring thread carries transcripts for other occurrences, so a
    # thread-wide bool would report ready for a stale/other recording (issue
    # #79). Fall back to thread-wide only when the chosen recording has no
    # callId to match against (older data), preserving prior behavior.
    if call_id:
        transcript_ready = any(
            ((m.get("eventDetail") or {}).get("callId") or "") == call_id
            for m in transcript_events
        )
    else:
        transcript_ready = bool(transcript_events)

    return {
        "recording_url": recording_url,
        "display_name": display_name,
        "transcript_ready": transcript_ready,
        "recording_date": recording_date,
        "occurrence_date": occurrence_date or None,
        "date_matches_occurrence": date_matches,
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
    result = await client.get(
        "/me/calendarView",
        params={
            "startDateTime": start_iso,
            "endDateTime": end_iso,
            "$orderby": "start/dateTime",
            "$top": "1",
        },
        headers=headers,
    )

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
        result = await client.get(
            "/me/calendarView",
            params={
                "startDateTime": start_iso,
                "endDateTime": end_iso,
                "$orderby": "start/dateTime asc",
                "$top": "50",
                "$filter": f"contains(subject, '{_escape_odata(subject)}')",
            },
            headers=headers,
        )

        events = result.get("value", [])
        if events:
            return events[0]
    except GraphAPIError:
        pass

    result = await client.get(
        "/me/calendarView",
        params={
            "startDateTime": start_iso,
            "endDateTime": end_iso,
            "$orderby": "start/dateTime asc",
            "$top": "50",
        },
        headers=headers,
    )

    subject_lower = subject.lower()
    for event in result.get("value", []):
        if subject_lower in event.get("subject", "").lower():
            return event

    return None
