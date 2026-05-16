"""whats_new composer — mail/calendar/teams since a given datetime."""

import asyncio

from ..formatters import (
    format_events_markdown,
    format_mail_summary_markdown,
    format_section_error,
    format_teams_activity_markdown,
)
from ..graph import GraphClient
from ..permissions import PermissionRegistry
from ._utils import NOISE_PATTERNS, _error_reason

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

    tasks = {}

    if scope in ("calendar", "all"):
        now_end = "2099-12-31T23:59:59"
        cal_params = {
            "startDateTime": since,
            "endDateTime": now_end,
            "$orderby": "start/dateTime",
            "$top": "20",
            "$select": "subject,start,end,location,attendees,organizer,isOnlineMeeting,onlineMeeting",
        }
        cal_headers = GraphClient.calendar_headers(timezone)
        tasks["calendar"] = client.get("/me/calendarView", params=cal_params, headers=cal_headers)

    mail_unavailable = permissions.check("Mail.Read")
    if scope in ("mail", "all") and not mail_unavailable:
        tasks["messages"] = client.get("/me/messages", params={
            "$filter": f"receivedDateTime ge {since}",
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
                relevant = [m for m in all_msgs if not _is_noise(m)]
                high_importance = [
                    {"subject": m.get("subject", "?"), "from": _sender_name(m)}
                    for m in relevant if m.get("importance") == "high"
                ]
                needs_attention = [
                    {"subject": m.get("subject", "?"), "from": _sender_name(m)}
                    for m in relevant[:5]
                ]
                sections.append(format_mail_summary_markdown(
                    unread_count=len(all_msgs),
                    relevant_count=len(relevant),
                    flagged_count=0,
                    high_importance=high_importance[:5],
                    needs_attention=needs_attention,
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
                        from_field = (preview.get("from") or {})
                        user_field = (from_field.get("user") or {})
                        display_name = user_field.get("displayName", "Unknown")
                        preview_msgs.append({
                            "from": {"user": {"displayName": display_name}},
                            "body": preview.get("body", {}),
                        })
                sections.append(format_teams_activity_markdown(preview_msgs))

    return "\n\n".join(sections) if sections else "Nothing new since that time."


def _is_noise(msg: dict) -> bool:
    from_addr = (msg.get("from") or {}).get("emailAddress", {}).get("address", "").lower()
    return any(p in from_addr for p in NOISE_PATTERNS)


def _sender_name(msg: dict) -> str:
    return (msg.get("from") or {}).get("emailAddress", {}).get("name", "Unknown")
