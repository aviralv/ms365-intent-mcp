"""my_day composer — calendar + mail + teams for a single day."""

import asyncio

from ..formatters import (
    _extract_event_links,
    _strip_teams_boilerplate,
    _strip_teams_html,
    format_events_markdown,
    format_mail_summary_markdown,
    format_section_error,
    format_teams_activity_markdown,
    graph_dt_to_aware_iso,
)
from ..graph import GraphClient
from ..permissions import PermissionRegistry
from ._utils import _build_mail_summary, _chat_sender, _error_reason


async def compose_my_day(
    client: GraphClient,
    permissions: PermissionRegistry,
    date: str,
    timezone: str,
    include_bodies: bool = False,
) -> tuple[dict, str]:
    start_iso = f"{date}T00:00:00"
    end_iso = f"{date}T23:59:59"

    tasks = {}

    select = "subject,start,end,location,attendees,organizer,isOnlineMeeting,onlineMeeting"
    if include_bodies:
        select += ",body"
    cal_params = {
        "startDateTime": start_iso,
        "endDateTime": end_iso,
        "$orderby": "start/dateTime",
        "$top": "30",
        "$select": select,
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
    events_raw: list[dict] = []
    unread_count = 0
    unread_msgs: list[dict] = []
    chats_raw: list[dict] = []

    # Calendar section
    cal_result = results.get("calendar")
    if isinstance(cal_result, BaseException):
        sections.append(format_section_error("Calendar", _error_reason(cal_result)))
    else:
        events_raw = cal_result.get("value", []) if cal_result else []
        sections.append(f"### Calendar\n{format_events_markdown(events_raw)}")

    # Mail section
    if mail_unavailable:
        sections.append(f"### Mail\n{mail_unavailable}")
    else:
        inbox_result = results.get("inbox_count")
        unread_result = results.get("unread")
        inbox_err = inbox_result if isinstance(inbox_result, BaseException) else None
        unread_err = unread_result if isinstance(unread_result, BaseException) else None
        first_err = inbox_err or unread_err
        if first_err is not None:
            sections.append(format_section_error("Mail", _error_reason(first_err)))
        else:
            unread_count = (
                inbox_result.get("unreadItemCount", 0)
                if isinstance(inbox_result, dict) else 0
            )
            unread_msgs = (
                unread_result.get("value", [])
                if isinstance(unread_result, dict) else []
            )
            summary = _build_mail_summary(unread_msgs)
            sections.append(format_mail_summary_markdown(
                unread_count=unread_count,
                relevant_count=summary["relevant_count"],
                flagged_count=0,
                high_importance=summary["high_importance"],
                needs_attention=summary["needs_attention"],
            ))

    # Teams section
    if teams_unavailable:
        sections.append(f"### Teams\n{teams_unavailable}")
    else:
        chats_result = results.get("chats")
        if isinstance(chats_result, BaseException):
            sections.append(format_section_error("Teams", _error_reason(chats_result)))
        else:
            chats_raw = chats_result.get("value", []) if chats_result else []
            preview_msgs = []
            for chat in chats_raw[:5]:
                preview = chat.get("lastMessagePreview")
                if preview:
                    preview_msgs.append({
                        "from": {"user": {"displayName": _chat_sender(preview)}},
                        "body": preview.get("body", {}),
                    })
            sections.append(format_teams_activity_markdown(preview_msgs))

    markdown = "\n\n".join(sections)

    # Build structured data dict
    event_list = []
    for e in events_raw:
        start_iso, start_tz = graph_dt_to_aware_iso(e.get("start", {}))
        end_iso, end_tz = graph_dt_to_aware_iso(e.get("end", {}))
        event_entry = {
            "subject": e.get("subject", ""),
            "start": start_iso,
            "end": end_iso,
            "start_timezone": start_tz,
            "end_timezone": end_tz,
            "location": e.get("location", {}).get("displayName") or None,
            "is_online_meeting": bool(e.get("isOnlineMeeting")),
        }
        if include_bodies:
            raw_body = (e.get("body") or {}).get("content", "")
            event_entry["links"] = _extract_event_links(raw_body, e)
            text = _strip_teams_html(raw_body, preserve_links=True) if raw_body else ""
            text = _strip_teams_boilerplate(text)
            event_entry["body"] = text[:2000] or None
        event_list.append(event_entry)

    mail_summary = _build_mail_summary(unread_msgs) if unread_msgs else {"relevant_count": 0, "high_importance": [], "needs_attention": [], "all_count": 0}
    teams_count = sum(
        1 for chat in chats_raw if chat.get("lastMessagePreview")
    )

    data = {
        "date": date,
        "events": event_list,
        "mail": {
            "unread_count": unread_count,
            "relevant_count": mail_summary.get("relevant_count", 0),
            "flagged_count": 0,
        },
        "teams": {
            "recent_message_count": teams_count,
        },
    }
    return data, markdown


