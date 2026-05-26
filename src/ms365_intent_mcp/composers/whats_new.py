"""whats_new composer — mail/calendar/teams since a given datetime."""

import asyncio
from datetime import datetime, timedelta, timezone as _tz

from ..formatters import (
    format_events_markdown,
    format_mail_summary_markdown,
    format_section_error,
    format_teams_activity_markdown,
)
from ..graph import GraphClient
from ..permissions import PermissionRegistry
from ._utils import _build_mail_summary, _chat_sender, _error_reason

_VALID_SCOPES = {"mail", "calendar", "teams", "all"}


async def compose_whats_new(
    client: GraphClient,
    permissions: PermissionRegistry,
    since: str,
    scope: str | None,
    timezone: str,
) -> str:
    scope = (scope or "all").lower()
    if scope not in _VALID_SCOPES:
        scope = "all"

    try:
        since_dt = datetime.fromisoformat(since.replace("Z", "+00:00"))
    except (ValueError, TypeError):
        return "❌ Invalid 'since' format. Use ISO datetime, e.g. '2026-05-14T00:00:00'."
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
        since_z = since if since.endswith("Z") else since + "Z"
        tasks["messages"] = client.get("/me/messages", params={
            "$filter": f"receivedDateTime ge {since_z}",
            "$select": "subject,from,receivedDateTime,importance",
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

    sections = []

    if "calendar" in results:
        cal_result = results["calendar"]
        if isinstance(cal_result, Exception):
            sections.append(format_section_error("Calendar", _error_reason(cal_result)))
        else:
            events = (cal_result or {}).get("value", [])
            sections.append(f"### Calendar\n{format_events_markdown(events)}")

    if scope in ("mail", "all"):
        if mail_unavailable:
            sections.append(f"### Mail\n{mail_unavailable}")
        elif "messages" in results:
            msgs_result = results["messages"]
            if isinstance(msgs_result, Exception):
                sections.append(format_section_error("Mail", _error_reason(msgs_result)))
            else:
                all_msgs = (msgs_result or {}).get("value", [])
                summary = _build_mail_summary(all_msgs)
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
            if isinstance(chats_result, Exception):
                sections.append(format_section_error("Teams", _error_reason(chats_result)))
            else:
                chats = (chats_result or {}).get("value", [])
                preview_msgs = []
                for chat in chats[:5]:
                    preview = chat.get("lastMessagePreview")
                    if preview:
                        preview_msgs.append({
                            "from": {"user": {"displayName": _chat_sender(preview)}},
                            "body": preview.get("body", {}),
                            "_chat_web_url": chat.get("webUrl", ""),
                        })
                sections.append(format_teams_activity_markdown(preview_msgs))

    return "\n\n".join(sections) if sections else "Nothing new since that time."


