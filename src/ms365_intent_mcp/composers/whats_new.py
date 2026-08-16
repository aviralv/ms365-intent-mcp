"""whats_new composer — mail/calendar/teams since a given datetime."""

import asyncio
from datetime import datetime, timedelta
from datetime import timezone as _tz

from ..formatters import (
    format_events_markdown,
    format_mail_summary_markdown,
    format_section_error,
    format_teams_activity_markdown,
    graph_dt_to_aware_iso,
)
from ..graph import GraphClient
from ..permissions import PermissionRegistry
from ._utils import _build_mail_summary, _chat_sender, _error_reason

_VALID_SCOPES = {"mail", "calendar", "teams", "all"}

# Per-chat message fetch is bounded to the most-recently-active chats so the
# extra Graph calls stay small (issue #67).
_TEAMS_CHAT_FANOUT = 5


def _parse_graph_dt(raw: str | None) -> datetime | None:
    """Parse a Graph ISO timestamp (handles trailing 'Z' and fractional seconds)."""
    if not raw:
        return None
    try:
        dt = datetime.fromisoformat(raw.replace("Z", "+00:00"))
    except (ValueError, TypeError):
        return None
    return dt.replace(tzinfo=_tz.utc) if dt.tzinfo is None else dt


async def _fetch_chat_window_messages(
    client: GraphClient, chat_id: str, since_dt: datetime
) -> list[dict]:
    """Fetch a chat's recent messages and keep those created at/after ``since_dt``.

    Graph's ``$filter`` support on ``/chats/{id}/messages`` is inconsistent, so we
    pull the most recent page and window-filter client-side. Returns newest-first,
    skipping system-event messages and empty-body entries (both are noise).
    """
    resp = await client.get(f"/me/chats/{chat_id}/messages", params={"$top": "20"})
    msgs = (resp or {}).get("value", [])
    in_window = []
    for m in msgs:
        created = _parse_graph_dt(m.get("createdDateTime"))
        if created is None or created < since_dt:
            continue
        content = (m.get("body") or {}).get("content", "").strip()
        # Skip call-started/member-added/etc. system events. Graph wraps them with
        # body '<systemEventMessage/>' regardless of messageType (which is the
        # unreliable 'unknownFutureValue', not 'systemEventMessage') — so key off
        # the body marker. Also skip genuinely empty bodies. (issue #67)
        if not content or content.startswith("<systemEventMessage"):
            continue
        in_window.append(m)
    return in_window


async def compose_whats_new(
    client: GraphClient,
    permissions: PermissionRegistry,
    since: str,
    scope: str | None,
    timezone: str,
) -> tuple[dict, str]:
    scope = (scope or "all").lower()
    if scope not in _VALID_SCOPES:
        scope = "all"

    try:
        since_dt = datetime.fromisoformat(since.replace("Z", "+00:00"))
    except (ValueError, TypeError):
        markdown = "❌ Invalid 'since' format. Use ISO datetime, e.g. '2026-05-14T00:00:00'."
        return {"since": since, "mail": [], "events": [], "teams": []}, markdown
    if since_dt.tzinfo is None:
        since_dt = since_dt.replace(tzinfo=_tz.utc)

    tasks = {}

    if scope in ("calendar", "all"):
        now = datetime.now(_tz.utc)
        max_end = since_dt + timedelta(days=14)
        cal_end = min(now + timedelta(days=7), max_end)
        cal_end_iso = cal_end.strftime("%Y-%m-%dT23:59:59")
        cal_params = {
            "startDateTime": since,
            "endDateTime": cal_end_iso,
            "$orderby": "start/dateTime",
            "$top": "20",
            "$select": "subject,start,end,location,attendees,organizer,isOnlineMeeting,onlineMeeting",
        }
        cal_headers = GraphClient.calendar_headers(timezone)
        tasks["calendar"] = client.get("/me/calendarView", params=cal_params, headers=cal_headers)

    mail_unavailable = permissions.check("Mail.Read")
    if scope in ("mail", "all") and not mail_unavailable:
        # Graph's $filter parser requires RFC-3339 with a single UTC marker.
        # Format from the parsed datetime so we don't have to guess whether the
        # caller-supplied string ends in 'Z', '+00:00', or nothing at all.
        since_z = since_dt.astimezone(_tz.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        tasks["messages"] = client.get("/me/messages", params={
            "$filter": f"receivedDateTime ge {since_z}",
            "$select": "id,subject,from,receivedDateTime,importance,webLink",
            "$orderby": "receivedDateTime desc",
            "$top": "25",
        })

    teams_unavailable = permissions.check("Chat.ReadWrite")
    if scope in ("teams", "all") and not teams_unavailable:
        tasks["chats"] = client.get("/me/chats", params={
            "$top": "10",
            "$expand": "lastMessagePreview",
            "$orderby": "lastMessagePreview/createdDateTime desc",
        })

    keys = list(tasks.keys())
    results_list = await asyncio.gather(*tasks.values(), return_exceptions=True)
    results = dict(zip(keys, results_list))

    # Issue #67: the chats list only carries lastMessagePreview (one message per
    # chat), which masks an inbound reply whenever the user's own message is the
    # most recent. Fetch each surfaced chat's in-window messages so every message
    # since `since_dt` is returned, not just the latest.
    chat_messages: dict[str, list[dict]] = {}
    chats_for_fetch: list[dict] = []
    if scope in ("teams", "all") and not teams_unavailable and "chats" in results:
        chats_result = results["chats"]
        if not isinstance(chats_result, BaseException):
            chats_for_fetch = [
                c for c in (chats_result or {}).get("value", [])[:_TEAMS_CHAT_FANOUT]
                if c.get("id")
            ]
            msg_results = await asyncio.gather(
                *(_fetch_chat_window_messages(client, c["id"], since_dt) for c in chats_for_fetch),
                return_exceptions=True,
            )
            for chat, mres in zip(chats_for_fetch, msg_results):
                if isinstance(mres, BaseException):
                    # Fall back to the preview so a transient per-chat failure
                    # never surfaces less than the old behavior did.
                    preview = chat.get("lastMessagePreview")
                    chat_messages[chat["id"]] = [preview] if preview else []
                else:
                    chat_messages[chat["id"]] = mres

    sections = []
    events_raw: list[dict] = []
    mail_raw: list[dict] = []

    if "calendar" in results:
        cal_result = results["calendar"]
        if isinstance(cal_result, BaseException):
            sections.append(format_section_error("Calendar", _error_reason(cal_result)))
        else:
            events_raw = (cal_result or {}).get("value", [])
            sections.append(f"### Calendar\n{format_events_markdown(events_raw)}")

    if scope in ("mail", "all"):
        if mail_unavailable:
            sections.append(f"### Mail\n{mail_unavailable}")
        elif "messages" in results:
            msgs_result = results["messages"]
            if isinstance(msgs_result, BaseException):
                sections.append(format_section_error("Mail", _error_reason(msgs_result)))
            else:
                mail_raw = (msgs_result or {}).get("value", [])
                summary = _build_mail_summary(mail_raw)
                sections.append(format_mail_summary_markdown(
                    unread_count=summary["all_count"],
                    relevant_count=summary["relevant_count"],
                    flagged_count=0,
                    high_importance=summary["high_importance"],
                    needs_attention=summary["needs_attention"],
                ))

    if scope in ("teams", "all"):
        if teams_unavailable:
            sections.append(f"### Teams\n{teams_unavailable}")
        elif "chats" in results:
            chats_result = results["chats"]
            if isinstance(chats_result, BaseException):
                sections.append(format_section_error("Teams", _error_reason(chats_result)))
            else:
                preview_msgs = []
                for chat in chats_for_fetch:
                    for msg in chat_messages.get(chat["id"], []):
                        preview_msgs.append({
                            "from": {"user": {"displayName": _chat_sender(msg)}},
                            "body": msg.get("body", {}),
                            "_chat_web_url": chat.get("webUrl", ""),
                        })
                sections.append(format_teams_activity_markdown(preview_msgs))

    markdown = "\n\n".join(sections) if sections else "Nothing new since that time."

    # Build structured data
    event_list = []
    for e in events_raw:
        start_iso, start_tz = graph_dt_to_aware_iso(e.get("start", {}))
        end_iso, end_tz = graph_dt_to_aware_iso(e.get("end", {}))
        event_list.append({
            "subject": e.get("subject", ""),
            "start": start_iso,
            "end": end_iso,
            "start_timezone": start_tz,
            "end_timezone": end_tz,
            "location": e.get("location", {}).get("displayName") or None,
            "is_online_meeting": bool(e.get("isOnlineMeeting")),
        })

    mail_list = []
    for m in mail_raw:
        mail_list.append({
            "subject": m.get("subject", ""),
            "sender": (m.get("from") or {}).get("emailAddress", {}).get("name", ""),
            "received": m.get("receivedDateTime", ""),
            "is_read": bool(m.get("isRead", False)),
            "importance": (m.get("importance") or "normal").lower(),
            "message_id": m.get("id") or None,
            "web_link": m.get("webLink") or None,
        })

    teams_list = []
    for chat in chats_for_fetch:
        for msg in chat_messages.get(chat["id"], []):
            teams_list.append({
                "chat_name": chat.get("topic"),
                "sender": _chat_sender(msg),
                "body_preview": (msg.get("body") or {}).get("content", "")[:200],
                "received": (msg.get("createdDateTime") or ""),
            })

    data = {
        "since": since,
        "mail": mail_list,
        "events": event_list,
        "teams": teams_list,
    }
    return data, markdown


