"""Markdown formatters for all tool responses."""

import re


def format_events_markdown(events: list[dict]) -> str:
    if not events:
        return "No events scheduled."
    lines = []
    for e in events:
        start = e.get("start", {}).get("dateTime", "")
        end = e.get("end", {}).get("dateTime", "")
        start_time = start[11:16] if len(start) > 16 else start
        end_time = end[11:16] if len(end) > 16 else end
        subject = e.get("subject", "(no subject)")
        location = e.get("location", {}).get("displayName", "")
        online = " 📹" if e.get("isOnlineMeeting") else ""
        attendee_names = [a["emailAddress"]["name"] for a in e.get("attendees", [])]
        attendees_str = f" — {', '.join(attendee_names[:5])}" if attendee_names else ""
        loc_str = f" | {location}" if location else ""
        lines.append(f"- **{start_time}–{end_time}** {subject}{online}{loc_str}{attendees_str}")
    return "\n".join(lines)


def format_event_detail_markdown(event: dict) -> str:
    subject = event.get("subject", "(no subject)")
    start = event.get("start", {}).get("dateTime", "")
    end = event.get("end", {}).get("dateTime", "")
    organizer = event.get("organizer", {}).get("emailAddress", {}).get("name", "Unknown")
    location = event.get("location", {}).get("displayName", "")
    body = event.get("body", {}).get("content", "")
    online = event.get("isOnlineMeeting", False)

    lines = [f"## {subject}"]
    lines.append(f"**When:** {start[:16]} → {end[11:16]}")
    lines.append(f"**Organizer:** {organizer}")
    if location:
        lines.append(f"**Location:** {location}")
    if online:
        join_url = (event.get("onlineMeeting") or {}).get("joinUrl", "")
        lines.append(f"**Teams meeting:** {join_url}" if join_url else "**Teams meeting**")

    attendees = event.get("attendees", [])
    if attendees:
        lines.append("\n**Attendees:**")
        for a in attendees[:20]:
            name = a.get("emailAddress", {}).get("name", "Unknown")
            status = a.get("status", {}).get("response", "none")
            emoji = {"accepted": "✅", "declined": "❌", "tentative": "❓"}.get(status, "⬜")
            lines.append(f"- {emoji} {name}")

    if body:
        text = re.sub(r"<[^>]+>", "", body).strip()
        if text:
            excerpt = text[:500]
            lines.append(f"\n**Body:**\n{excerpt}")

    return "\n".join(lines)


def format_mail_summary_markdown(
    unread_count: int,
    relevant_count: int,
    flagged_count: int,
    high_importance: list[dict],
    needs_attention: list[dict],
) -> str:
    lines = ["### Mail"]
    lines.append(f"- **{unread_count}** unread ({relevant_count} relevant, noise filtered)")
    if flagged_count:
        lines.append(f"- **{flagged_count}** flagged for follow-up")
    if high_importance:
        lines.append("- **High importance:**")
        for e in high_importance[:5]:
            lines.append(f"  - {e['subject']} (from {e['from']})")
    if needs_attention:
        lines.append("- **Needs attention:**")
        for e in needs_attention[:5]:
            lines.append(f"  - {e['subject']} (from {e['from']})")
    return "\n".join(lines)


def format_draft_created_markdown(draft: dict) -> str:
    subject = draft.get("subject", "(no subject)")
    recipients = draft.get("toRecipients", [])
    to_names = [r.get("emailAddress", {}).get("name", "?") for r in recipients]
    lines = [
        "✅ Draft created",
        f"**Subject:** {subject}",
        f"**To:** {', '.join(to_names)}",
        "",
        "Draft saved to your Drafts folder. Open Outlook to review and send.",
    ]
    return "\n".join(lines)


def format_event_created_markdown(event: dict) -> str:
    subject = event.get("subject", "(no subject)")
    start = event.get("start", {}).get("dateTime", "")[:16]
    lines = [
        "✅ Event created",
        f"**Subject:** {subject}",
        f"**When:** {start}",
    ]
    if event.get("isOnlineMeeting"):
        join_url = (event.get("onlineMeeting") or {}).get("joinUrl", "")
        if join_url:
            lines.append(f"**Teams link:** {join_url}")
    return "\n".join(lines)


def format_section_error(section: str, reason: str) -> str:
    return f"### {section}\n⚠️  {section} unavailable — {reason}."


def format_teams_activity_markdown(messages: list[dict]) -> str:
    if not messages:
        return "No recent Teams activity."
    lines = ["### Teams"]
    for msg in messages[:5]:
        sender = msg.get("from", {}).get("user", {}).get("displayName", "Unknown")
        body = msg.get("body", {}).get("content", "")
        text = re.sub(r"<[^>]+>", "", body).strip()[:100]
        lines.append(f"- **{sender}:** {text}")
    return "\n".join(lines)
