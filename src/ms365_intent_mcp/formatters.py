"""Markdown formatters for all tool responses."""

import re


def _strip_teams_html(body: str) -> str:
    """Strip Teams HTML for plain-text rendering.

    Replaces <at id="...">@Name</at> mention tags with their inner text so
    mention-only messages don't render as empty. Then strips all remaining tags.
    """
    if not body:
        return ""
    body = re.sub(r"<at\b[^>]*>(.*?)</at>", r"\1", body, flags=re.IGNORECASE | re.DOTALL)
    body = re.sub(r"<[^>]+>", "", body)
    return body.strip()


def _format_event_time_range(start: dict, end: dict) -> str:
    """Format a Graph dateTimeTimeZone pair as 'HH:MM–HH:MM TZ'.

    Graph returns event times as {dateTime: "...", timeZone: "..."} pairs (e.g.
    `{"dateTime": "2026-06-02T07:45:00.0000000", "timeZone": "UTC"}`). With
    `Prefer: outlook.timezone` set, the dateTime is a naive string with no Z;
    the timezone lives in the sibling field. Always include it so consumers
    can't mistake UTC for local time.

    Falls back to bare 'HH:MM–HH:MM' if no timeZone is present (defensive —
    Graph always sends one).
    """
    s = start.get("dateTime", "")
    e = end.get("dateTime", "")
    tz = start.get("timeZone") or end.get("timeZone") or ""
    s_hm = s[11:16] if len(s) >= 16 else s
    e_hm = e[11:16] if len(e) >= 16 else e
    suffix = f" {tz}" if tz else ""
    return f"{s_hm}–{e_hm}{suffix}"


def _format_event_datetime(dt: dict) -> str:
    """Format a single Graph dateTimeTimeZone object as 'YYYY-MM-DDTHH:MM TZ'.

    Used in single-event detail views where the date matters (a meeting can
    span dates, and the detail view shows only one timestamp at a time).

    Returns empty string for {} so call sites can pass through naturally.
    """
    s = dt.get("dateTime", "")
    tz = dt.get("timeZone") or ""
    s_fmt = s[:16] if len(s) >= 16 else s
    suffix = f" {tz}" if tz else ""
    if not s_fmt and not suffix:
        return ""
    return f"{s_fmt}{suffix}"


def _format_offset_datetime(ts: str | None) -> str:
    """Format a Graph dateTimeOffset ISO string as 'YYYY-MM-DDTHH:MM UTC'.

    Graph dateTimeOffset fields (chatMessage.createdDateTime,
    message.receivedDateTime, driveItem.lastModifiedDateTime, etc.) are always
    UTC and always end in Z, sometimes with milliseconds (e.g.
    '2026-05-29T10:00:00.035Z'). The Prefer: outlook.timezone header does NOT
    apply to these fields — they are always UTC.

    Slicing to [:16] takes 'YYYY-MM-DDTHH:MM' regardless of millisecond
    presence, then we append the literal ' UTC' so callers can't mistake the
    naive-looking string for local time.

    Accepts None for callers passing `data.get("createdDateTime")` without
    a default. Returns empty string for None / empty input. Returns the
    value untouched (without UTC suffix) for strings shorter than 16 chars
    to avoid mislabelling junk.
    """
    if not ts:
        return ""
    if len(ts) < 16:
        return ts
    return f"{ts[:16]} UTC"


def format_events_markdown(events: list[dict]) -> str:
    if not events:
        return "No events scheduled."
    lines = []
    for e in events:
        time_range = _format_event_time_range(e.get("start", {}), e.get("end", {}))
        subject = e.get("subject", "(no subject)")
        location = e.get("location", {}).get("displayName", "")
        online = " 📹" if e.get("isOnlineMeeting") else ""
        attendee_names = [a["emailAddress"]["name"] for a in e.get("attendees", [])]
        attendees_str = f" — {', '.join(attendee_names[:5])}" if attendee_names else ""
        loc_str = f" | {location}" if location else ""
        lines.append(f"- **{time_range}** {subject}{online}{loc_str}{attendees_str}")
    return "\n".join(lines)


def format_event_detail_markdown(event: dict) -> str:
    subject = event.get("subject", "(no subject)")
    start = event.get("start", {})
    end = event.get("end", {})
    organizer = event.get("organizer", {}).get("emailAddress", {}).get("name", "Unknown")
    location = event.get("location", {}).get("displayName", "")
    body = event.get("body", {}).get("content", "")
    online = event.get("isOnlineMeeting", False)

    lines = [f"## {subject}"]
    start_fmt = _format_event_datetime(start)
    end_fmt = _format_event_datetime(end)
    lines.append(f"**When:** {start_fmt} → {end_fmt}")
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
        text = _strip_teams_html(body)
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
    when = _format_event_datetime(event.get("start", {}))
    lines = [
        "✅ Event created",
        f"**Subject:** {subject}",
        f"**When:** {when}",
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
        text = _strip_teams_html(body)
        if len(text) > 500:
            text = text[:500] + "…"
        web_url = msg.get("_chat_web_url", "")
        link = f" — [open chat]({web_url})" if web_url else ""
        lines.append(f"- **{sender}:** {text}{link}")
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
            text = _strip_teams_html(body)[:80]
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
            text = _strip_teams_html(body)[:80]
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
        start = slot.get("start", {})
        end = slot.get("end", {})
        confidence = s.get("confidence", 0)
        start_fmt = _format_event_datetime(start) or "?"
        end_dt = end.get("dateTime", "")
        end_tz = end.get("timeZone") or start.get("timeZone") or ""
        end_hm = end_dt[11:16] if len(end_dt) >= 16 else end_dt or "?"
        end_fmt = f"{end_hm} {end_tz}".strip()
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
        text = _strip_teams_html(body)[:300]
        sender = data.get("from", {}).get("user", {}).get("displayName", "?")
        created = _format_offset_datetime(data.get("createdDateTime", ""))
        return f"### Teams Message\n**From:** {sender}  |  **At:** {created}\n\n{text}"
    elif url_type == "meeting":
        return format_event_detail_markdown(data)
    elif url_type == "chat_thread":
        return _format_chat_thread(data)
    elif url_type == "sharepoint_page":
        if data.get("_page_found"):
            title = data.get("title", "") or data.get("name", "?")
            site_name = data.get("_site_name", "")
            url = data.get("webUrl", "")
            modified = data.get("lastModifiedDateTime", "")[:10]
            lines = [f"### SharePoint Page: {title}"]
            if site_name:
                lines.append(f"**Site:** {site_name}")
            if modified:
                lines.append(f"**Modified:** {modified}")
            if url:
                lines.append(f"**URL:** {url}")
            return "\n".join(lines)
        else:
            name = data.get("displayName", data.get("name", "?"))
            url = data.get("webUrl", "")
            lines = [f"### SharePoint Site: {name}"]
            if url:
                lines.append(f"**URL:** {url}")
            lines.append("_(Page content unavailable — showing site info)_")
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


def _format_chat_thread(data: dict) -> str:
    chat = data.get("chat") or {}
    entries = data.get("entries") or []
    meeting = data.get("meeting")
    chat_error = data.get("_chat_error")
    messages_error = data.get("_messages_error")

    lines: list[str] = []

    # Header: topic, else fall back to member names, else generic.
    members = chat.get("members") or []
    member_names = [m.get("displayName", "?") for m in members if m.get("displayName")]
    topic = chat.get("topic")
    if topic:
        header = topic
    elif member_names:
        header = ", ".join(member_names[:3])
    else:
        header = "Teams Chat"
    lines.append(f"### Teams Chat: {header}")

    chat_type = chat.get("chatType")
    if chat_type:
        lines.append(f"**Type:** {chat_type}")

    if member_names:
        shown = member_names[:6]
        more = len(member_names) - len(shown)
        members_str = ", ".join(shown)
        if more > 0:
            members_str += f" + {more} more"
        lines.append(f"**Members:** {members_str}")

    if chat_error:
        lines.append(f"⚠️  Chat metadata unavailable — {chat_error}.")
    if messages_error:
        lines.append(f"⚠️  Messages unavailable — {messages_error}.")

    if meeting:
        m_subject = meeting.get("subject", "(no subject)")
        m_start = _format_event_datetime(meeting.get("start", {}) or {})
        end_dt_obj = meeting.get("end", {}) or {}
        end_dt = end_dt_obj.get("dateTime", "")
        end_tz = end_dt_obj.get("timeZone") or (meeting.get("start", {}) or {}).get("timeZone") or ""
        end_hm = end_dt[11:16] if len(end_dt) >= 16 else end_dt
        m_end = f"{end_hm} {end_tz}".strip()
        m_organizer = ((meeting.get("organizer", {}) or {}).get("emailAddress", {}) or {}).get("name", "Unknown")
        lines.append("")
        lines.append("**Meeting context:**")
        lines.append(f"- {m_subject}")
        lines.append(f"- {m_start} → {m_end}")
        lines.append(f"- Organizer: {m_organizer}")

    if entries:
        lines.append("")
        lines.append("**Recent activity:**")
        for entry in entries[:25]:
            lines.append(_format_chat_entry(entry))

    return "\n".join(lines)


def _format_chat_entry(entry: dict) -> str:
    """Render one entry from the chat_thread `entries` list.

    Timestamps are rendered as 'YYYY-MM-DD HH:MM UTC' for full clarity (chats
    can span months/years; UTC label avoids timezone confusion). Calls show
    date once + time range when same day; both dates when crossing midnight.
    """
    raw_ts = (entry.get("ts") or "")[:16]  # 'YYYY-MM-DDTHH:MM'
    ts_display = raw_ts.replace("T", " ")  # 'YYYY-MM-DD HH:MM'
    ts_with_tz = f"{ts_display} UTC" if ts_display else ts_display
    kind = entry.get("kind")

    if kind == "message":
        sender = entry.get("sender") or "Unknown"
        if entry.get("is_body_empty"):
            return f"- **{sender}** ({ts_with_tz}): _(no text)_"
        return f"- **{sender}** ({ts_with_tz}): {entry.get('body', '')}"

    if kind == "call":
        raw_end = (entry.get("end_ts") or "")[:16]
        start_date, start_time = raw_ts[:10], raw_ts[11:16]
        end_date, end_time = raw_end[:10], raw_end[11:16]
        if not raw_end or raw_end == raw_ts:
            # Single-event call: just the start, with UTC
            time_range = f"{start_date} {start_time} UTC"
        elif start_date == end_date:
            # Same-day: date once, time range, UTC on both sides
            time_range = f"{start_date} {start_time} UTC–{end_time} UTC"
        else:
            # Cross-day: full both ends, UTC on both
            time_range = f"{start_date} {start_time} UTC → {end_date} {end_time} UTC"
        duration = entry.get("duration")
        if duration:
            time_range += f", {duration}"
        initiator = entry.get("initiator")
        header = "📞 **Call**"
        if initiator:
            header = f"📞 **Call started by {initiator}**"
        parts = [f"{header} ({time_range})"]
        if entry.get("recording_url"):
            parts.append(f"[recording]({entry['recording_url']})")
        if entry.get("transcript_ready"):
            parts.append("transcript ready")
        return "- " + " — ".join(parts)

    if kind == "event":
        return f"- ⚙️ {entry.get('summary', 'system event')} ({ts_with_tz})"

    return f"- _(unknown entry: {kind})_"
