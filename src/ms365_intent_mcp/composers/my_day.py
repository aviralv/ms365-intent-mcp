"""my_day composer — calendar + mail + teams for a single day."""

import asyncio

from ..formatters import (
    format_events_markdown,
    format_mail_summary_markdown,
    format_section_error,
    format_teams_activity_markdown,
)
from ..graph import GraphClient
from ..permissions import PermissionRegistry
from ._utils import _chat_sender, _error_reason, _is_noise, _sender_name


async def compose_my_day(
    client: GraphClient,
    permissions: PermissionRegistry,
    date: str,
    timezone: str,
) -> str:
    start_iso = f"{date}T00:00:00"
    end_iso = f"{date}T23:59:59"

    tasks = {}

    cal_params = {
        "startDateTime": start_iso,
        "endDateTime": end_iso,
        "$orderby": "start/dateTime",
        "$top": "30",
        "$select": "subject,start,end,location,attendees,organizer,isOnlineMeeting,onlineMeeting",
    }
    cal_headers = GraphClient.calendar_headers(timezone)
    tasks["calendar"] = client.get("/me/calendarView", params=cal_params, headers=cal_headers)

    mail_unavailable = permissions.check("Mail.Read")
    if not mail_unavailable:
        tasks["inbox_count"] = client.get(
            "/me/mailFolders/inbox", params={"$select": "unreadItemCount"}
        )
        tasks["unread"] = client.get("/me/mailFolders/inbox/messages", params={
            "$filter": "isRead eq false",
            "$select": "subject,from,receivedDateTime,importance",
            "$orderby": "receivedDateTime desc",
            "$top": "20",
        })

    teams_unavailable = permissions.check("Chat.ReadWrite")
    if not teams_unavailable:
        tasks["chats"] = client.get("/me/chats", params={
            "$top": "10",
            "$expand": "lastMessagePreview",
            "$orderby": "lastMessagePreview/createdDateTime desc",
        })

    keys = list(tasks.keys())
    results_list = await asyncio.gather(*tasks.values(), return_exceptions=True)
    results = dict(zip(keys, results_list))

    sections = []

    # Calendar section
    cal_result = results.get("calendar")
    if isinstance(cal_result, Exception):
        sections.append(format_section_error("Calendar", _error_reason(cal_result)))
    else:
        events = cal_result.get("value", []) if cal_result else []
        sections.append(f"### Calendar\n{format_events_markdown(events)}")

    # Mail section
    if mail_unavailable:
        sections.append(f"### Mail\n{mail_unavailable}")
    else:
        inbox_result = results.get("inbox_count")
        unread_result = results.get("unread")
        if isinstance(inbox_result, Exception) or isinstance(unread_result, Exception):
            err = inbox_result if isinstance(inbox_result, Exception) else unread_result
            sections.append(format_section_error("Mail", _error_reason(err)))
        else:
            unread_count = inbox_result.get("unreadItemCount", 0) if inbox_result else 0
            unread_msgs = unread_result.get("value", []) if unread_result else []
            relevant = [m for m in unread_msgs if not _is_noise(m)]
            high_importance = [
                {"subject": m.get("subject", "?"), "from": _sender_name(m)}
                for m in relevant if m.get("importance") == "high"
            ]
            needs_attention = [
                {"subject": m.get("subject", "?"), "from": _sender_name(m)}
                for m in relevant[:5]
            ]
            sections.append(format_mail_summary_markdown(
                unread_count=unread_count,
                relevant_count=len(relevant),
                flagged_count=0,
                high_importance=high_importance[:5],
                needs_attention=needs_attention,
            ))

    # Teams section
    if teams_unavailable:
        sections.append(f"### Teams\n{teams_unavailable}")
    else:
        chats_result = results.get("chats")
        if isinstance(chats_result, Exception):
            sections.append(format_section_error("Teams", _error_reason(chats_result)))
        else:
            chats = chats_result.get("value", []) if chats_result else []
            preview_msgs = []
            for chat in chats[:5]:
                preview = chat.get("lastMessagePreview")
                if preview:
                    preview_msgs.append({
                        "from": {"user": {"displayName": _chat_sender(preview)}},
                        "body": preview.get("body", {}),
                    })
            sections.append(format_teams_activity_markdown(preview_msgs))

    return "\n\n".join(sections)


