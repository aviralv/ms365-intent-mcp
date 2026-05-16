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


def format_people_markdown(
    query: str,
    people: list[dict],
    recent_emails: list[dict],
    recent_chat: dict | None,
) -> str:
    if not people:
        return f"### People\nNo results for '{query}'."
    lines = [f"### People — {query}"]
    person = people[0]
    name = person.get("displayName", "Unknown")
    emails = person.get("emailAddresses", [])
    email_addr = emails[0].get("address", "") if emails else ""
    job_title = person.get("jobTitle", "")
    lines.append(f"**{name}**" + (f" — {job_title}" if job_title else ""))
    if email_addr:
        lines.append(f"📧 {email_addr}")
    if len(people) > 1:
        others = [p.get("displayName", "?") for p in people[1:4]]
        lines.append(f"Also matched: {', '.join(others)}")
    if recent_emails:
        lines.append("\n**Recent mail:**")
        for m in recent_emails[:3]:
            subject = m.get("subject", "(no subject)")
            sender = m.get("from", {}).get("emailAddress", {}).get("name", "?")
            lines.append(f"- {subject} (from {sender})")
    if recent_chat:
        preview = (recent_chat.get("lastMessagePreview") or {})
        body = preview.get("body", {}).get("content", "") if preview else ""
        if body:
            text = re.sub(r"<[^>]+>", "", body).strip()[:80]
            lines.append(f"\n**Recent Teams chat:** {text}")
    return "\n".join(lines)


def format_search_results_markdown(query: str, hits: list[dict]) -> str:
    if not hits:
        return f"### Search — '{query}'\nNo results found."
    lines = [f"### Search — '{query}' ({len(hits)} result{'s' if len(hits) != 1 else ''})"]
    for hit in hits[:10]:
        resource = hit.get("resource", {})
        odata_type = resource.get("@odata.type", "")
        if "message" in odata_type:
            subject = resource.get("subject", "(no subject)")
            sender = resource.get("from", {}).get("emailAddress", {}).get("name", "?")
            preview = resource.get("bodyPreview", "")[:80]
            lines.append(f"- **[Mail]** {subject} — *from {sender}*")
            if preview:
                lines.append(f"  {preview}")
        elif "driveItem" in odata_type:
            name = resource.get("name", "?")
            web_url = resource.get("webUrl", "")
            lines.append(f"- **[File]** {name}" + (f" — {web_url}" if web_url else ""))
        elif "chatMessage" in odata_type:
            body = resource.get("body", {}).get("content", "")
            text = re.sub(r"<[^>]+>", "", body).strip()[:80]
            lines.append(f"- **[Teams]** {text}")
        elif "listItem" in odata_type:
            fields = resource.get("fields", {})
            title = fields.get("Title", resource.get("name", "?"))
            lines.append(f"- **[SharePoint]** {title}")
        else:
            lines.append(f"- {resource.get('name', resource.get('subject', '?'))}")
    return "\n".join(lines)


def format_meeting_times_markdown(suggestions: list[dict]) -> str:
    if not suggestions:
        return "### Schedule\nNo available time slots found for those constraints."
    lines = ["### Available Meeting Times"]
    for i, s in enumerate(suggestions[:5], 1):
        slot = s.get("meetingTimeSlot", {})
        start = slot.get("start", {}).get("dateTime", "")
        end = slot.get("end", {}).get("dateTime", "")
        confidence = s.get("confidence", 0)
        start_fmt = start[:16] if start else "?"
        end_fmt = end[11:16] if end else "?"
        lines.append(f"{i}. **{start_fmt} – {end_fmt}** ({confidence:.0f}% confidence)")
        unavailable = [
            a.get("attendee", {}).get("emailAddress", {}).get("name", "?")
            for a in s.get("attendeeAvailability", [])
            if a.get("availability") not in ("free", "unknown")
        ]
        if unavailable:
            lines.append(f"   Conflict: {', '.join(unavailable)}")
    return "\n".join(lines)


def format_resolved_content_markdown(url_type: str, data: dict) -> str:
    if url_type == "email":
        subject = data.get("subject", "(no subject)")
        sender = data.get("from", {}).get("emailAddress", {}).get("name", "?")
        received = data.get("receivedDateTime", "")[:10]
        preview = data.get("bodyPreview", "")[:200]
        lines = [f"### Email: {subject}", f"**From:** {sender}  |  **Received:** {received}"]
        if preview:
            lines.append(f"\n{preview}")
        return "\n".join(lines)
    elif url_type in ("channel_message", "chat_message"):
        body = data.get("body", {}).get("content", "")
        text = re.sub(r"<[^>]+>", "", body).strip()[:300]
        sender = data.get("from", {}).get("user", {}).get("displayName", "?")
        created = data.get("createdDateTime", "")[:16]
        return f"### Teams Message\n**From:** {sender}  |  **At:** {created}\n\n{text}"
    elif url_type == "meeting":
        return format_event_detail_markdown(data)
    elif url_type == "sharepoint_page":
        name = data.get("displayName", data.get("name", "?"))
        url = data.get("webUrl", "")
        lines = [f"### SharePoint Page: {name}"]
        if url:
            lines.append(f"**URL:** {url}")
        return "\n".join(lines)
    elif url_type in ("onedrive_file", "onedrive_share_link"):
        name = data.get("name", "?")
        size = data.get("size", 0)
        url = data.get("webUrl", "")
        size_kb = size // 1024 if size else 0
        lines = [f"### File: {name}", f"**Size:** {size_kb} KB"]
        if url:
            lines.append(f"**URL:** {url}")
        return "\n".join(lines)
    else:
        return f"### Resolved\n```\n{data}\n```"
