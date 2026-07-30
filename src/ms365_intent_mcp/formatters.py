"""Markdown formatters for all tool responses."""

import datetime as _dt
import logging
import re
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from .windows_tz import windows_to_iana

logger = logging.getLogger(__name__)


_MAX_EMAIL_BODY_BYTES = 200 * 1024


def _teams_open_chat_link(chat_url: str) -> str:
    """Render a trailing ' — [open chat](url)' suffix, or '' when no url.

    Shared by the Teams-activity, people, and search-hit formatters so the
    link label and separator live in one place.
    """
    return f" — [open chat]({chat_url})" if chat_url else ""


def _truncate_email_body(text: str) -> str:
    """Cap an email body at ~200 KB to protect the LLM context.

    Truncation is by UTF-8 byte length, not character count — a 300 KB body of
    mostly ASCII is worth capping, but the same character count of CJK is 3× as
    many bytes. Cuts on a codepoint boundary and appends a visible marker.
    Real email bodies observed in the wild are ~4 KB; this cap is safety, not
    routine trimming.
    """
    if not text:
        return ""
    encoded = text.encode("utf-8")
    if len(encoded) <= _MAX_EMAIL_BODY_BYTES:
        return text
    truncated = encoded[:_MAX_EMAIL_BODY_BYTES].decode("utf-8", errors="ignore")
    return truncated + f"\n\n…[truncated at {_MAX_EMAIL_BODY_BYTES // 1024} KB]"


def _strip_teams_html(body: str, preserve_links: bool = False) -> str:
    """Strip Teams HTML for plain-text rendering.

    Replaces <at id="...">@Name</at> mention tags with their inner text so
    mention-only messages don't render as empty, then strips all remaining tags.

    When `preserve_links=True`, anchors are first converted to markdown links so
    URLs (to other conversations, pages, files) survive — Teams bodies carry
    links as <a href="URL">text</a> and dropping the href makes the referent
    invisible. When the display text equals the href (or is empty), the bare URL
    is emitted to avoid [url](url). Default is False (plain text) because most
    callers truncate the result with a hard slice, which would cut markdown link
    syntax mid-token; opt in only where the full, untruncated body is rendered.
    """
    if not body:
        return ""
    if preserve_links:
        body = re.sub(
            r'<a\b[^>]*\bhref="([^"]*)"[^>]*>(.*?)</a>',
            _anchor_to_markdown,
            body,
            flags=re.IGNORECASE | re.DOTALL,
        )
    body = re.sub(r"<at\b[^>]*>(.*?)</at>", r"\1", body, flags=re.IGNORECASE | re.DOTALL)
    body = re.sub(r"<[^>]+>", "", body)
    return body.strip()


def _anchor_to_markdown(match: re.Match) -> str:
    """Render an <a href> match as markdown '[text](url)', or bare url when the
    inner text is empty or already equals the href."""
    url = match.group(1).strip()
    text = re.sub(r"<[^>]+>", "", match.group(2)).strip()
    if not url:
        return text
    if not text or text == url:
        return url
    return f"[{text}]({url})"


def graph_dt_to_aware_iso(dt: dict) -> tuple[str | None, str | None]:
    """Convert a Graph {dateTime, timeZone} pair to (offset-aware ISO, tz name).

    Graph returns naive wall-clock strings under `Prefer: outlook.timezone`; the
    zone lives in the sibling field. This makes the structured value
    self-describing so consumers can't mistake it for local time.

    - Empty/{} pair                → (None, None)
    - Date-only (all-day, len 10)  → (date_string, tz_name)   [no offset — a date
                                       has no instant]
    - Timed + resolvable zone      → (localized offset ISO, tz_name)
    - Timed + unresolvable zone    → (naive ISO, tz_name) + logged warning

    fold=0 is used for DST-ambiguous wall-clocks; zoneinfo does not raise.
    """
    raw = (dt or {}).get("dateTime")
    if not raw:
        return None, None
    tz_name = (dt or {}).get("timeZone")
    s = raw.strip()

    # All-day events arrive as bare dates — pass through, no instant to localize.
    if len(s) == 10:
        return s, tz_name

    # Strip Graph's 7-digit fractional seconds so fromisoformat accepts it.
    core = s[:19] if len(s) >= 19 else s
    try:
        naive = _dt.datetime.fromisoformat(core)
    except ValueError:
        logger.warning("graph_dt_to_aware_iso: unparseable dateTime %r", raw)
        return core, tz_name

    zone = None
    if tz_name:
        if tz_name == "UTC":
            zone = _dt.timezone.utc
        else:
            iana = windows_to_iana(tz_name) or tz_name
            try:
                zone = ZoneInfo(iana)
            except (ZoneInfoNotFoundError, ValueError):
                zone = None

    if zone is None:
        if tz_name:
            logger.warning(
                "graph_dt_to_aware_iso: could not resolve timezone %r; "
                "emitting naive timestamp", tz_name,
            )
        return core, tz_name

    return naive.replace(tzinfo=zone, fold=0).isoformat(), tz_name


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

    recording = event.get("_recording")
    if recording:
        lines.append("")
        lines.append("**Recording:**")
        display_name = recording.get("display_name")
        if display_name:
            lines.append(f"- name: {display_name}")
        url = recording.get("recording_url")
        if url:
            lines.append(f"- link: {url}")
        vroom = recording.get("vroom_url")
        if vroom:
            lines.append(f"- vroom_url: `{vroom}`")
        drive_id = recording.get("drive_id")
        if drive_id:
            lines.append(f"- drive_id: `{drive_id}`")
        drive_item_id = recording.get("drive_item_id")
        if drive_item_id:
            lines.append(f"- drive_item_id: `{drive_item_id}`")
        owner = recording.get("owner_upn")
        if owner:
            lines.append(f"- owner: `{owner}`")
        if recording.get("transcript_ready"):
            lines.append("- transcript: ready")

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


def format_event_forwarded_markdown(to_names: list[str], comment: str | None) -> str:
    lines = [
        "✅ Meeting invite forwarded",
        f"**To:** {', '.join(to_names)}",
    ]
    if comment:
        lines.append(f"**Comment:** {comment}")
    lines.append("")
    lines.append("The invite was sent immediately. The organizer is notified and the recipient added to their copy.")
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
        link = _teams_open_chat_link(web_url)
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
        chat_url = recent_chat.get("webUrl", "")
        link = _teams_open_chat_link(chat_url)
        if body:
            text = _strip_teams_html(body)[:80]
            lines.append(f"\n**Recent Teams chat:** {text}{link}")
    return "\n".join(lines)


def format_search_results_markdown(query: str, hits: list[dict]) -> str:
    if not hits:
        return f"### Search — '{query}'\nNo results found."
    lines = [f"### Search — '{query}' ({len(hits)} result{'s' if len(hits) != 1 else ''})"]
    for hit in hits[:10]:
        resource = hit.get("resource", {})
        odata_type = resource.get("@odata.type", "")
        if "chatMessage" in odata_type:
            sender = (resource.get("from") or {}).get("user", {}).get("displayName", "?")
            body_html = (resource.get("body") or {}).get("content", "")
            text = _strip_teams_html(body_html)[:200]
            link = _teams_open_chat_link(resource.get("chat_url", ""))
            lines.append(f"- **[Teams]** {sender}: {text}{link}")
        elif "message" in odata_type:
            subject = resource.get("subject", "(no subject)")
            sender = (resource.get("from") or {}).get("emailAddress", {}).get("name", "?")
            preview = resource.get("bodyPreview") or ""
            web_link = resource.get("webLink") or ""
            lines.append(f"- **[Mail]** {subject} — *from {sender}*")
            if preview:
                lines.append(f"  {preview}")
            if web_link:
                # Follow-up hint: resolve(url=<webLink>) fetches the full body.
                lines.append(f"  🔗 {web_link}")
        elif "driveItem" in odata_type:
            name = resource.get("name", "?")
            web_url = resource.get("webUrl", "")
            lines.append(f"- **[File]** {name}" + (f" — {web_url}" if web_url else ""))
        elif "listItem" in odata_type:
            fields = resource.get("fields") or {}
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


def format_resolved_content_markdown(url_type: str, data: dict, chat_url: str = "") -> str:
    if url_type == "email":
        subject = data.get("subject", "(no subject)")
        sender = data.get("from", {}).get("emailAddress", {}).get("name", "?")
        received = data.get("receivedDateTime", "")[:10]
        body = data.get("body") or {}
        body_content = body.get("content", "")
        body_type = body.get("contentType", "")
        # Prefer full body (server returns text when Prefer header is honored);
        # strip HTML if the server ignored the Prefer and returned html anyway.
        # Fall back to bodyPreview only when body is entirely missing.
        if body_content:
            rendered = body_content if body_type == "text" else _strip_teams_html(body_content)
        else:
            rendered = (data.get("bodyPreview") or "")
        rendered = _truncate_email_body(rendered)
        lines = [f"### Email: {subject}", f"**From:** {sender}  |  **Received:** {received}"]
        if rendered:
            lines.append("")
            lines.append(rendered)
        attachments = data.get("_attachments") or []
        if attachments:
            lines.append("")
            lines.append("**Attachments:**")
            for a in attachments:
                lines.append(_format_email_attachment_line(a))
        att_error = data.get("_attachments_error")
        if att_error:
            lines.append(f"⚠️ Some attachments could not be retrieved: {att_error}")
        return "\n".join(lines)
    elif url_type in ("channel_message", "chat_message"):
        body = data.get("body", {}).get("content", "")
        text = _strip_teams_html(body)[:300]
        sender = data.get("from", {}).get("user", {}).get("displayName", "?")
        created = _format_offset_datetime(data.get("createdDateTime", ""))
        link = _teams_open_chat_link(chat_url)
        return f"### Teams Message\n**From:** {sender}  |  **At:** {created}{link}\n\n{text}"
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
            main = f"- **{sender}** ({ts_with_tz}): _(no text)_"
        else:
            main = f"- **{sender}** ({ts_with_tz}): {entry.get('body', '')}"
        sub_lines: list[str] = []
        reply = entry.get("reply_to")
        if reply:
            r_sender = reply.get("sender") or "Unknown"
            r_preview = reply.get("preview") or ""
            sub_lines.append(f"  - ↩️ replying to **{r_sender}**: {r_preview}")
        sub_lines.extend(_format_attachment_links(entry.get("attachments")))
        if sub_lines:
            return "\n".join([main] + sub_lines)
        return main

    if kind == "call":
        raw_end = (entry.get("end_ts") or "")[:16]
        start_date, start_time = raw_ts[:10], raw_ts[11:16]
        end_date, end_time = raw_end[:10], raw_end[11:16]
        if not raw_ts:
            # No timestamp: empty time_range, no orphan UTC, no double space
            time_range = ""
        elif not raw_end or raw_end == raw_ts:
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
        main_line = "- " + " — ".join(parts)
        # Structured drive metadata (from /shares/ enrichment) as sub-bullets.
        # These are what a caller like ferret-transcripts needs to download
        # directly, without name-search guessing.
        detail_lines = _format_recording_details(entry)
        if detail_lines:
            return "\n".join([main_line] + detail_lines)
        return main_line

    if kind == "event":
        return f"- ⚙️ {entry.get('summary', 'system event')} ({ts_with_tz})"

    return f"- _(unknown entry: {kind})_"


def _format_attachment_links(attachments: list[dict] | None) -> list[str]:
    """Render shared-file/recording attachments as 📎 link sub-bullets.

    Returns [] when there are no attachments. Each entry is a {name, url} dict
    from _extract_reference_attachments — a fetchable SharePoint link that would
    otherwise read as unrecoverable pasted media (issue #36)."""
    if not attachments:
        return []
    lines = []
    for att in attachments:
        url = att.get("url")
        if not url:
            continue
        name = att.get("name") or url
        lines.append(f"  - 📎 [{name}]({url})")
    return lines


def _human_size(n: float) -> str:
    for unit in ("B", "KB", "MB"):
        if n < 1024:
            return f"{int(n)} {unit}" if unit == "B" else f"{n:.1f} {unit}"
        n /= 1024
    return f"{n:.1f} GB"


def _format_email_attachment_line(a: dict) -> str:
    """One 📎 line for an email attachment: name · size · inline/file · path|note."""
    parts = [a.get("name") or "(unnamed)"]
    size = a.get("size") or 0
    if size:
        parts.append(_human_size(size))
    parts.append("inline" if a.get("is_inline") else "file")
    if a.get("local_path"):
        parts.append(f"saved: `{a['local_path']}`")
    elif a.get("note"):
        parts.append(a["note"])
    return "📎 " + " · ".join(parts)


def _format_recording_details(entry: dict) -> list[str]:
    """Render drive_id / drive_item_id / vroom_url as sub-bullets when present.

    Returns [] when the entry has no enriched fields (recording was surfaced
    URL-only, or /shares/ lookup failed). Callers that just want the link
    keep seeing the same one-line entry; callers that want ready-to-download
    metadata get the extra fields."""
    vroom = entry.get("vroom_url")
    drive_id = entry.get("drive_id")
    drive_item_id = entry.get("drive_item_id")
    owner = entry.get("owner_upn")
    if not (vroom or drive_id or drive_item_id):
        return []
    lines = []
    if vroom:
        lines.append(f"  - vroom_url: `{vroom}`")
    if drive_id:
        lines.append(f"  - drive_id: `{drive_id}`")
    if drive_item_id:
        lines.append(f"  - drive_item_id: `{drive_item_id}`")
    if owner:
        lines.append(f"  - owner: `{owner}`")
    return lines
