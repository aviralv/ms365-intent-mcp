"""Tests for markdown formatters."""

import logging

from ms365_intent_mcp.formatters import (
    format_events_markdown,
    format_event_detail_markdown,
    format_mail_summary_markdown,
    format_draft_created_markdown,
    format_event_created_markdown,
    format_section_error,
    format_teams_activity_markdown,
    format_people_markdown,
    format_search_results_markdown,
    format_meeting_times_markdown,
    format_resolved_content_markdown,
    graph_dt_to_aware_iso,
)


class TestFormatEventsMarkdown:
    def test_empty_list(self):
        result = format_events_markdown([])
        assert "No events" in result

    def test_single_event(self):
        events = [{
            "subject": "Standup",
            "start": {"dateTime": "2026-05-15T09:00:00", "timeZone": "UTC"},
            "end": {"dateTime": "2026-05-15T09:30:00", "timeZone": "UTC"},
            "location": {"displayName": "Room A"},
            "isOnlineMeeting": True,
            "attendees": [{"emailAddress": {"name": "Bob"}}],
            "organizer": {"emailAddress": {"name": "Alice"}},
        }]
        result = format_events_markdown(events)
        assert "Standup" in result
        assert "09:00–09:30 UTC" in result
        assert "Room A" in result

    def test_named_iana_timezone_appears_in_output(self):
        events = [{
            "subject": "Standup",
            "start": {"dateTime": "2026-05-15T09:00:00", "timeZone": "Europe/Berlin"},
            "end": {"dateTime": "2026-05-15T09:30:00", "timeZone": "Europe/Berlin"},
            "location": {"displayName": ""},
            "isOnlineMeeting": False,
            "attendees": [],
            "organizer": {"emailAddress": {"name": "X"}},
        }]
        result = format_events_markdown(events)
        assert "Europe/Berlin" in result

    def test_event_no_timezone_falls_back_gracefully(self):
        events = [{
            "subject": "Standup",
            "start": {"dateTime": "2026-05-15T09:00:00"},
            "end": {"dateTime": "2026-05-15T09:30:00"},
            "location": {"displayName": ""},
            "isOnlineMeeting": False,
            "attendees": [],
            "organizer": {"emailAddress": {"name": "X"}},
        }]
        result = format_events_markdown(events)
        assert "Standup" in result
        assert "09:00–09:30" in result
        assert "None" not in result

    def test_multiple_events_ordered(self):
        events = [
            {
                "subject": "First",
                "start": {"dateTime": "2026-05-15T09:00:00"},
                "end": {"dateTime": "2026-05-15T09:30:00"},
                "location": {"displayName": ""},
                "isOnlineMeeting": False,
                "attendees": [],
                "organizer": {"emailAddress": {"name": "X"}},
            },
            {
                "subject": "Second",
                "start": {"dateTime": "2026-05-15T10:00:00"},
                "end": {"dateTime": "2026-05-15T11:00:00"},
                "location": {"displayName": ""},
                "isOnlineMeeting": False,
                "attendees": [],
                "organizer": {"emailAddress": {"name": "Y"}},
            },
        ]
        result = format_events_markdown(events)
        assert result.index("First") < result.index("Second")


class TestFormatEventDetailMarkdown:
    def test_includes_all_fields(self):
        event = {
            "subject": "Team Sync",
            "start": {"dateTime": "2026-05-15T14:00:00", "timeZone": "UTC"},
            "end": {"dateTime": "2026-05-15T14:30:00", "timeZone": "UTC"},
            "location": {"displayName": "Room B"},
            "isOnlineMeeting": True,
            "onlineMeeting": {"joinUrl": "https://teams.microsoft.com/l/meetup-join/123"},
            "organizer": {"emailAddress": {"name": "Alice"}},
            "attendees": [
                {"emailAddress": {"name": "Bob"}, "status": {"response": "accepted"}},
                {"emailAddress": {"name": "Carol"}, "status": {"response": "declined"}},
            ],
            "body": {"content": "<p>Agenda: Q2 review</p>", "contentType": "html"},
        }
        result = format_event_detail_markdown(event)
        assert "Team Sync" in result
        assert "Alice" in result
        assert "Room B" in result
        assert "teams.microsoft.com" in result
        assert "Bob" in result
        assert "Agenda" in result or "Q2 review" in result
        assert "2026-05-15T14:00 UTC" in result
        assert "2026-05-15T14:30 UTC" in result


    def test_cross_date_event_renders_both_dates(self):
        event = {
            "subject": "Late-night ops",
            "start": {"dateTime": "2026-05-15T23:30:00", "timeZone": "UTC"},
            "end": {"dateTime": "2026-05-16T00:15:00", "timeZone": "UTC"},
            "location": {"displayName": ""},
            "isOnlineMeeting": False,
            "onlineMeeting": None,
            "organizer": {"emailAddress": {"name": "Alice"}},
            "attendees": [],
            "body": {"content": ""},
        }
        result = format_event_detail_markdown(event)
        assert "Late-night ops" in result
        assert "2026-05-15T23:30 UTC" in result
        assert "2026-05-16T00:15 UTC" in result

    def test_end_missing_timezone_renders_gracefully(self):
        """The new code drops the end.tz → start.tz fallback. Verify the
        helper handles missing end.timeZone without crashing or producing
        a 'None' artifact in the output. With Prefer: outlook.timezone set,
        Graph always returns matching tz on both sides — but this test
        guards the regression path explicitly."""
        event = {
            "subject": "Standup",
            "start": {"dateTime": "2026-05-15T14:00:00", "timeZone": "UTC"},
            "end": {"dateTime": "2026-05-15T14:30:00"},  # no timeZone
            "location": {"displayName": ""},
            "isOnlineMeeting": False,
            "onlineMeeting": None,
            "organizer": {"emailAddress": {"name": "Alice"}},
            "attendees": [],
            "body": {"content": ""},
        }
        result = format_event_detail_markdown(event)
        # Start side has UTC; end side renders bare date+time without UTC.
        assert "2026-05-15T14:00 UTC" in result
        assert "2026-05-15T14:30" in result
        # No crash; no orphan "None" artifact.
        assert "None" not in result

    def test_recording_block_rendered_when_recording_present(self):
        """When compose_meeting attaches _recording, format_event_detail_markdown
        renders a Recording block with structured fields."""
        event = {
            "subject": "Team Sync",
            "start": {"dateTime": "2026-05-15T14:00:00", "timeZone": "UTC"},
            "end": {"dateTime": "2026-05-15T14:30:00", "timeZone": "UTC"},
            "location": {"displayName": ""},
            "isOnlineMeeting": True,
            "onlineMeeting": {"joinUrl": "https://x"},
            "organizer": {"emailAddress": {"name": "Alice"}},
            "attendees": [],
            "body": {"content": ""},
            "_recording": {
                "recording_url": "https://sap-my.sharepoint.com/:v:/p/alice/IQBw",
                "display_name": "Team Sync-recording.mp4",
                "transcript_ready": True,
                "drive_id": "b!xxx",
                "drive_item_id": "01ABC",
                "owner_upn": "alice",
                "vroom_url": "https://sap-my.sharepoint.com/personal/alice/_api/v2.0/drives/b!xxx/items/01ABC",
            },
        }
        result = format_event_detail_markdown(event)
        assert "**Recording:**" in result
        assert "Team Sync-recording.mp4" in result
        assert "https://sap-my.sharepoint.com/:v:/p/alice/IQBw" in result
        assert "b!xxx" in result
        assert "01ABC" in result
        assert "alice" in result
        assert "transcript: ready" in result

    def test_no_recording_block_when_recording_absent(self):
        """Events without _recording render as before — no blank Recording block."""
        event = {
            "subject": "Team Sync",
            "start": {"dateTime": "2026-05-15T14:00:00", "timeZone": "UTC"},
            "end": {"dateTime": "2026-05-15T14:30:00", "timeZone": "UTC"},
            "location": {"displayName": ""},
            "isOnlineMeeting": True,
            "onlineMeeting": {"joinUrl": "https://x"},
            "organizer": {"emailAddress": {"name": "Alice"}},
            "attendees": [],
            "body": {"content": ""},
        }
        result = format_event_detail_markdown(event)
        assert "Recording" not in result

    def test_recording_block_with_only_url_and_link(self):
        """Degraded case: /shares/ 403'd, only URL survived enrichment.
        Should still render the link, just no drive fields."""
        event = {
            "subject": "Team Sync",
            "start": {"dateTime": "2026-05-15T14:00:00", "timeZone": "UTC"},
            "end": {"dateTime": "2026-05-15T14:30:00", "timeZone": "UTC"},
            "location": {"displayName": ""},
            "isOnlineMeeting": True,
            "onlineMeeting": {"joinUrl": "https://x"},
            "organizer": {"emailAddress": {"name": "Alice"}},
            "attendees": [],
            "body": {"content": ""},
            "_recording": {
                "recording_url": "https://sap-my.sharepoint.com/:v:/p/other/XYZ",
                "display_name": "recording.mp4",
                "transcript_ready": False,
            },
        }
        result = format_event_detail_markdown(event)
        assert "**Recording:**" in result
        assert "recording.mp4" in result
        assert "https://sap-my.sharepoint.com/:v:/p/other/XYZ" in result
        # No drive fields — they weren't enrichable.
        assert "drive_id" not in result
        assert "vroom_url" not in result


class TestFormatSectionError:
    def test_timeout_message(self):
        result = format_section_error("Mail", "timed out")
        assert "⚠️" in result
        assert "Mail" in result
        assert "timed out" in result


class TestFormatDraftCreated:
    def test_includes_subject(self):
        draft = {
            "subject": "Follow up",
            "id": "draft-123",
            "toRecipients": [{"emailAddress": {"name": "Bob", "address": "bob@example.com"}}],
        }
        result = format_draft_created_markdown(draft)
        assert "Follow up" in result
        assert "Bob" in result


class TestFormatEventCreated:
    def test_includes_subject_and_time(self):
        event = {
            "subject": "New Meeting",
            "start": {"dateTime": "2026-05-16T10:00:00", "timeZone": "UTC"},
            "end": {"dateTime": "2026-05-16T10:30:00", "timeZone": "UTC"},
            "isOnlineMeeting": False,
            "onlineMeeting": None,
        }
        result = format_event_created_markdown(event)
        assert "New Meeting" in result
        assert "2026-05-16T10:00 UTC" in result


class TestFormatTeamsActivity:
    def test_empty_messages(self):
        result = format_teams_activity_markdown([])
        assert "No recent" in result

    def test_with_messages(self):
        msgs = [{
            "from": {"user": {"displayName": "Dave"}},
            "body": {"content": "Hey team!", "contentType": "text"},
        }]
        result = format_teams_activity_markdown(msgs)
        assert "Dave" in result
        assert "Hey team!" in result


class TestFormatMailSummary:
    def test_basic_summary(self):
        result = format_mail_summary_markdown(
            unread_count=10,
            relevant_count=7,
            flagged_count=2,
            high_importance=[{"subject": "Urgent", "from": "Boss"}],
            needs_attention=[{"subject": "Review PR", "from": "Dev"}],
        )
        assert "10" in result
        assert "7" in result
        assert "Urgent" in result


class TestFormatPeopleMarkdown:
    def test_empty_list(self):
        result = format_people_markdown("alice", [], [], None)
        assert "alice" in result.lower() or "No results" in result

    def test_with_person_and_mail(self):
        people = [{"displayName": "Alice Smith", "emailAddresses": [{"address": "alice@example.com"}]}]
        emails = [{"subject": "Project update", "from": {"emailAddress": {"name": "Bob"}}, "receivedDateTime": "2026-05-15T10:00:00Z"}]
        result = format_people_markdown("alice", people, emails, None)
        assert "Alice Smith" in result
        assert "Project update" in result

    def test_with_teams_chat(self):
        people = [{"displayName": "Alice Smith", "emailAddresses": [{"address": "alice@example.com"}]}]
        chat = {"id": "19:abc", "chatType": "oneOnOne", "lastMessagePreview": {"body": {"content": "Hey!"}}}
        result = format_people_markdown("alice", people, [], chat)
        assert "Alice Smith" in result


class TestFormatSearchResultsMarkdown:
    def test_empty_results(self):
        result = format_search_results_markdown("budget", [])
        assert "No results" in result or "budget" in result.lower()

    def test_with_results(self):
        hits = [
            {
                "resource": {
                    "@odata.type": "#microsoft.graph.message",
                    "subject": "Q2 Budget Review",
                    "from": {"emailAddress": {"name": "Finance", "address": "finance@example.com"}},
                    "receivedDateTime": "2026-05-10T09:00:00Z",
                    "bodyPreview": "Please find attached the Q2 budget...",
                }
            }
        ]
        result = format_search_results_markdown("budget", hits)
        assert "Q2 Budget Review" in result
        assert "Finance" in result

    def test_email_with_null_from(self):
        hits = [
            {
                "resource": {
                    "@odata.type": "#microsoft.graph.message",
                    "subject": "Automated report",
                    "from": None,
                    "bodyPreview": "System-generated notice",
                }
            }
        ]
        result = format_search_results_markdown("report", hits)
        assert "Automated report" in result

    def test_email_with_null_body_preview(self):
        hits = [
            {
                "resource": {
                    "@odata.type": "#microsoft.graph.message",
                    "subject": "Empty preview",
                    "from": {"emailAddress": {"name": "Sender"}},
                    "bodyPreview": None,
                }
            }
        ]
        result = format_search_results_markdown("empty", hits)
        assert "Empty preview" in result

    def test_listitem_with_null_fields(self):
        hits = [
            {
                "resource": {
                    "@odata.type": "#microsoft.graph.listItem",
                    "fields": None,
                    "name": "fallback-name",
                }
            }
        ]
        result = format_search_results_markdown("q", hits)
        assert "fallback-name" in result


    def test_email_weblink_rendered_when_present(self):
        hits = [
            {
                "resource": {
                    "@odata.type": "#microsoft.graph.message",
                    "subject": "Full body needed",
                    "from": {"emailAddress": {"name": "Alessia"}},
                    "bodyPreview": "First 255 chars of preview",
                    "webLink": "https://outlook.office365.com/owa/?ItemID=AAAA",
                }
            }
        ]
        result = format_search_results_markdown("q", hits)
        assert "Full body needed" in result
        # Follow-up hint present so the caller can resolve() the full body.
        assert "https://outlook.office365.com/owa/?ItemID=AAAA" in result

    def test_email_weblink_omitted_when_absent(self):
        """No webLink line, no crash — Graph doesn't guarantee webLink."""
        hits = [
            {
                "resource": {
                    "@odata.type": "#microsoft.graph.message",
                    "subject": "No link here",
                    "from": {"emailAddress": {"name": "Sender"}},
                    "bodyPreview": "Body preview text",
                    # webLink intentionally absent
                }
            }
        ]
        result = format_search_results_markdown("q", hits)
        assert "No link here" in result
        assert "🔗" not in result

    def test_email_bodypreview_not_truncated_to_80(self):
        """bodyPreview is already capped at 255 by Graph; don't re-cap at 80."""
        long_preview = "A" * 200
        hits = [
            {
                "resource": {
                    "@odata.type": "#microsoft.graph.message",
                    "subject": "Long preview",
                    "from": {"emailAddress": {"name": "S"}},
                    "bodyPreview": long_preview,
                }
            }
        ]
        result = format_search_results_markdown("q", hits)
        assert long_preview in result


class TestFormatMeetingTimesMarkdown:
    def test_empty_suggestions(self):
        result = format_meeting_times_markdown([])
        assert "No available" in result or "No slots" in result

    def test_with_suggestions(self):
        suggestions = [
            {
                "meetingTimeSlot": {
                    "start": {"dateTime": "2026-05-20T10:00:00", "timeZone": "UTC"},
                    "end": {"dateTime": "2026-05-20T10:30:00", "timeZone": "UTC"},
                },
                "confidence": 100.0,
                "attendeeAvailability": [],
            }
        ]
        result = format_meeting_times_markdown(suggestions)
        assert "2026-05-20T10:00 UTC" in result
        assert "10:30 UTC" in result
        assert "100" in result


class TestFormatResolvedContentMarkdown:
    def test_email_type(self):
        data = {
            "subject": "Hello",
            "from": {"emailAddress": {"name": "Bob", "address": "bob@example.com"}},
            "receivedDateTime": "2026-05-15T08:00:00Z",
            "bodyPreview": "Hi there",
        }
        result = format_resolved_content_markdown("email", data)
        assert "Hello" in result
        assert "Bob" in result

    def test_email_renders_full_text_body(self):
        """When body.contentType=='text', render body.content as-is (no strip)."""
        full = "Line one.\n\nLine two with detail.\n\nLine three."
        data = {
            "subject": "Details",
            "from": {"emailAddress": {"name": "Alessia"}},
            "receivedDateTime": "2026-07-01T09:00:00Z",
            "body": {"contentType": "text", "content": full},
            "bodyPreview": "Line one. Line two",
        }
        result = format_resolved_content_markdown("email", data)
        assert "Line two with detail" in result
        assert "Line three" in result

    def test_email_strips_html_if_server_ignored_prefer(self):
        html = "<html><body><p>Hello <b>world</b>.</p></body></html>"
        data = {
            "subject": "HTML fallback",
            "from": {"emailAddress": {"name": "Alessia"}},
            "receivedDateTime": "2026-07-01T09:00:00Z",
            "body": {"contentType": "html", "content": html},
            "bodyPreview": "Hello world.",
        }
        result = format_resolved_content_markdown("email", data)
        assert "Hello world" in result
        # No raw tags leaking through.
        assert "<b>" not in result
        assert "<html>" not in result

    def test_email_falls_back_to_body_preview_when_body_missing(self):
        data = {
            "subject": "Preview only",
            "from": {"emailAddress": {"name": "Sender"}},
            "receivedDateTime": "2026-07-01T09:00:00Z",
            "bodyPreview": "Just a preview.",
            # body key intentionally absent
        }
        result = format_resolved_content_markdown("email", data)
        assert "Just a preview." in result

    def test_email_truncates_at_200kb(self):
        """Pathological body caps at 200 KB with a visible marker."""
        # 300 KB of ASCII.
        oversize = "x" * (300 * 1024)
        data = {
            "subject": "Huge",
            "from": {"emailAddress": {"name": "S"}},
            "receivedDateTime": "2026-07-01T09:00:00Z",
            "body": {"contentType": "text", "content": oversize},
        }
        result = format_resolved_content_markdown("email", data)
        assert "truncated" in result.lower()
        # Result body portion must be smaller than the input.
        assert len(result) < len(oversize)

    def test_sharepoint_page_type(self):
        data = {"displayName": "Project Overview", "webUrl": "https://contoso.sharepoint.com/sites/proj"}
        result = format_resolved_content_markdown("sharepoint_page", data)
        assert "Project Overview" in result

    def test_onedrive_file_type(self):
        data = {"name": "report.xlsx", "size": 20480, "webUrl": "https://contoso-my.sharepoint.com/files/1"}
        result = format_resolved_content_markdown("onedrive_file", data)
        assert "report.xlsx" in result

    def test_chat_message_type_includes_utc_label(self):
        data = {
            "body": {"content": "Hello team"},
            "from": {"user": {"displayName": "Alice"}},
            "createdDateTime": "2026-05-29T10:00:00.035Z",
        }
        result = format_resolved_content_markdown("chat_message", data)
        assert "Alice" in result
        assert "Hello team" in result
        assert "2026-05-29T10:00 UTC" in result

    def test_chat_message_renders_open_chat_link_when_url_given(self):
        data = {
            "body": {"content": "Hello team"},
            "from": {"user": {"displayName": "Alice"}},
            "createdDateTime": "2026-05-29T10:00:00.035Z",
        }
        url = "https://teams.microsoft.com/l/chat/19:x@unq.gbl.spaces"
        result = format_resolved_content_markdown("chat_message", data, chat_url=url)
        assert "open chat" in result
        assert url in result


class TestFormatResolvedChatThread:
    def _data(self, **overrides):
        base = {
            "chat": {
                "topic": "Project Sync",
                "chatType": "meeting",
                "webUrl": "https://teams.microsoft.com/l/chat/19:abc@thread.v2/conversations",
                "members": [
                    {"displayName": "Alice"},
                    {"displayName": "Bob"},
                ],
            },
            "entries": [
                {"kind": "message",
                 "ts": "2026-05-26T10:05:00Z",
                 "sender": "Alice",
                 "body": "Hello world",
                 "is_body_empty": False},
                {"kind": "message",
                 "ts": "2026-05-26T10:00:00Z",
                 "sender": "Bob",
                 "body": "Hi",
                 "is_body_empty": False},
            ],
            "meeting": {
                "subject": "Project Sync",
                "start": {"dateTime": "2026-05-26T10:00:00", "timeZone": "UTC"},
                "end": {"dateTime": "2026-05-26T10:30:00", "timeZone": "UTC"},
                "organizer": {"emailAddress": {"name": "Alice"}},
            },
            "_chat_error": None,
            "_messages_error": None,
        }
        base.update(overrides)
        return base

    def test_renders_topic_and_messages(self):
        result = format_resolved_content_markdown("chat_thread", self._data())
        assert "Project Sync" in result
        assert "Alice" in result
        assert "Hello world" in result

    def test_renders_meeting_block_when_present(self):
        result = format_resolved_content_markdown("chat_thread", self._data())
        assert "Meeting" in result
        assert "10:00" in result
        assert "UTC" in result

    def test_omits_meeting_block_when_absent(self):
        result = format_resolved_content_markdown(
            "chat_thread", self._data(meeting=None)
        )
        assert "Meeting" not in result

    def test_falls_back_to_member_names_when_topic_null(self):
        data = self._data(chat={
            "topic": None,
            "chatType": "oneOnOne",
            "webUrl": "https://teams.microsoft.com/l/chat/19:dm@unq.gbl.spaces/conversations",
            "members": [
                {"displayName": "Avi"},
                {"displayName": "Alice"},
            ],
        })
        result = format_resolved_content_markdown("chat_thread", data)
        assert "Avi" in result
        assert "Alice" in result

    def test_caps_member_list_at_six_with_more_indicator(self):
        members = [{"displayName": f"User{i}"} for i in range(10)]
        data = self._data(chat={
            "topic": "Big group",
            "chatType": "group",
            "webUrl": "https://teams.microsoft.com/l/chat/19:abc@thread.v2/conversations",
            "members": members,
        })
        result = format_resolved_content_markdown("chat_thread", data)
        assert "User0" in result
        assert "User5" in result
        assert "+ 4 more" in result
        assert "User6" not in result

    def test_truncates_long_message_with_ellipsis(self):
        pre_truncated_body = "x" * 500 + "…"
        data = self._data(entries=[
            {"kind": "message",
             "ts": "2026-05-26T10:00:00Z",
             "sender": "Alice",
             "body": pre_truncated_body,
             "is_body_empty": False},
        ])
        result = format_resolved_content_markdown("chat_thread", data)
        assert "x" * 500 in result
        assert "…" in result
        assert "x" * 501 not in result

    def test_renders_chat_error_warning(self):
        data = self._data(chat=None, _chat_error="rate limited")
        result = format_resolved_content_markdown("chat_thread", data)
        assert "⚠️" in result
        assert "rate limited" in result

    def test_renders_messages_error_warning(self):
        data = self._data(entries=[], _messages_error="Microsoft service error")
        result = format_resolved_content_markdown("chat_thread", data)
        assert "⚠️" in result
        assert "Microsoft service error" in result

    def test_empty_message_body_renders_placeholder(self):
        data = self._data(entries=[
            {"kind": "message",
             "ts": "2026-05-26T10:00:00Z",
             "sender": "Alice",
             "body": "",
             "is_body_empty": True},
        ])
        result = format_resolved_content_markdown("chat_thread", data)
        assert "Alice" in result
        assert "_(no text)_" in result



class TestFormatTeamsActivityTruncation:
    def test_short_message_renders_intact(self):
        msgs = [{
            "from": {"user": {"displayName": "Alice"}},
            "body": {"content": "Hello"},
        }]
        result = format_teams_activity_markdown(msgs)
        assert "Hello" in result
        assert "…" not in result

    def test_message_at_500_chars_no_ellipsis(self):
        body = "x" * 500
        msgs = [{
            "from": {"user": {"displayName": "Alice"}},
            "body": {"content": body},
        }]
        result = format_teams_activity_markdown(msgs)
        assert body in result
        assert "…" not in result

    def test_message_above_500_chars_truncated_with_ellipsis(self):
        body = "x" * 600
        msgs = [{
            "from": {"user": {"displayName": "Alice"}},
            "body": {"content": body},
        }]
        result = format_teams_activity_markdown(msgs)
        assert "x" * 500 in result
        assert "…" in result
        assert "x" * 501 not in result

    def test_chat_web_url_renders_as_open_chat_link(self):
        msgs = [{
            "from": {"user": {"displayName": "Alice"}},
            "body": {"content": "Hello"},
            "_chat_web_url": "https://teams.microsoft.com/l/chat/19:abc@thread.v2/conversations",
        }]
        result = format_teams_activity_markdown(msgs)
        assert "[open chat]" in result
        assert "19:abc@thread.v2" in result

    def test_no_chat_web_url_no_link_appended(self):
        msgs = [{
            "from": {"user": {"displayName": "Alice"}},
            "body": {"content": "Hello"},
        }]
        result = format_teams_activity_markdown(msgs)
        assert "[open chat]" not in result


class TestStripTeamsHtml:
    def test_strips_generic_html_tags(self):
        from ms365_intent_mcp.formatters import _strip_teams_html
        assert _strip_teams_html("<p>Hello <b>world</b></p>") == "Hello world"

    def test_extracts_at_mention_inner_text(self):
        from ms365_intent_mcp.formatters import _strip_teams_html
        body = '<at id="b1b7e2b5-...">@Avi</at> please review'
        assert _strip_teams_html(body) == "@Avi please review"

    def test_at_mention_only_message_preserves_name(self):
        from ms365_intent_mcp.formatters import _strip_teams_html
        body = '<at id="123">@Avi</at>'
        assert _strip_teams_html(body) == "@Avi"

    def test_at_mention_inside_other_tags(self):
        from ms365_intent_mcp.formatters import _strip_teams_html
        body = '<p><at id="123">@Avi</at> hi</p>'
        assert _strip_teams_html(body) == "@Avi hi"

    def test_returns_empty_for_pure_html(self):
        from ms365_intent_mcp.formatters import _strip_teams_html
        assert _strip_teams_html("<systemEventMessage/>") == ""

    def test_empty_input(self):
        from ms365_intent_mcp.formatters import _strip_teams_html
        assert _strip_teams_html("") == ""

    def test_no_html_passthrough(self):
        from ms365_intent_mcp.formatters import _strip_teams_html
        assert _strip_teams_html("plain text") == "plain text"

    def test_strips_surrounding_whitespace(self):
        from ms365_intent_mcp.formatters import _strip_teams_html
        assert _strip_teams_html("  <p>hi</p>  ") == "hi"

    def test_anchor_preserved_as_markdown_link(self):
        from ms365_intent_mcp.formatters import _strip_teams_html
        body = 'See <a href="https://example.com/doc">the doc</a> now'
        assert _strip_teams_html(body, preserve_links=True) == "See [the doc](https://example.com/doc) now"

    def test_anchor_dropped_by_default(self):
        """Default (preserve_links=False) keeps only the anchor text — callers
        that hard-truncate the result must not receive markdown link syntax."""
        from ms365_intent_mcp.formatters import _strip_teams_html
        body = 'See <a href="https://example.com/doc">the doc</a> now'
        assert _strip_teams_html(body) == "See the doc now"

    def test_anchor_with_url_as_text_renders_bare_url(self):
        """When display text equals the href, don't produce [url](url)."""
        from ms365_intent_mcp.formatters import _strip_teams_html
        body = '<a href="https://example.com/x">https://example.com/x</a>'
        assert _strip_teams_html(body, preserve_links=True) == "https://example.com/x"

    def test_anchor_empty_text_renders_bare_url(self):
        from ms365_intent_mcp.formatters import _strip_teams_html
        body = '<a href="https://example.com/x"></a>'
        assert _strip_teams_html(body, preserve_links=True) == "https://example.com/x"

    def test_teams_conversation_link_preserved(self):
        from ms365_intent_mcp.formatters import _strip_teams_html
        url = "https://teams.microsoft.com/l/message/19:abc@thread.v2/123"
        body = f'ref: <a href="{url}">this thread</a>'
        assert _strip_teams_html(body, preserve_links=True) == f"ref: [this thread]({url})"

    def test_multiple_anchors_preserved(self):
        from ms365_intent_mcp.formatters import _strip_teams_html
        body = '<a href="https://a.com">A</a> and <a href="https://b.com">B</a>'
        assert _strip_teams_html(body, preserve_links=True) == "[A](https://a.com) and [B](https://b.com)"


class TestFormatChatEntry:
    def test_message_basic(self):
        from ms365_intent_mcp.formatters import _format_chat_entry
        entry = {
            "kind": "message",
            "ts": "2026-05-29T10:00:00Z",
            "sender": "Alice",
            "body": "Hello",
            "is_body_empty": False,
        }
        line = _format_chat_entry(entry)
        assert "Alice" in line
        assert "Hello" in line
        # Bug 4: render full date+time with space (not 'T') for readability; UTC label for clarity
        assert "2026-05-29 10:00 UTC" in line
        assert "_(no text)_" not in line

    def test_message_empty_body_renders_no_text(self):
        from ms365_intent_mcp.formatters import _format_chat_entry
        entry = {
            "kind": "message",
            "ts": "2026-05-29T10:00:00Z",
            "sender": "Alice",
            "body": "",
            "is_body_empty": True,
        }
        line = _format_chat_entry(entry)
        assert "_(no text)_" in line
        assert "_(deleted)_" not in line

    def test_message_with_reference_attachment_renders_link(self):
        from ms365_intent_mcp.formatters import _format_chat_entry
        entry = {
            "kind": "message",
            "ts": "2026-07-15T10:25:00Z",
            "sender": "Alice",
            "body": "",
            "is_body_empty": True,
            "attachments": [
                {"name": "Meeting Recording.mp4", "url": "https://x/rec.mp4"}
            ],
        }
        line = _format_chat_entry(entry)
        assert "📎 [Meeting Recording.mp4](https://x/rec.mp4)" in line
        assert line.split("\n")[0].startswith("- **Alice**")

    def test_message_with_text_and_attachment_renders_both(self):
        from ms365_intent_mcp.formatters import _format_chat_entry
        entry = {
            "kind": "message",
            "ts": "2026-07-15T10:25:00Z",
            "sender": "Alice",
            "body": "Here is the recording",
            "is_body_empty": False,
            "attachments": [
                {"name": "rec.mp4", "url": "https://x/rec.mp4"}
            ],
        }
        line = _format_chat_entry(entry)
        assert "Here is the recording" in line
        assert "📎 [rec.mp4](https://x/rec.mp4)" in line

    def test_message_without_attachments_no_paperclip(self):
        from ms365_intent_mcp.formatters import _format_chat_entry
        entry = {
            "kind": "message",
            "ts": "2026-07-15T10:25:00Z",
            "sender": "Alice",
            "body": "hi",
            "is_body_empty": False,
        }
        line = _format_chat_entry(entry)
        assert "📎" not in line

    def test_message_with_reply_context_renders_quote(self):
        from ms365_intent_mcp.formatters import _format_chat_entry
        entry = {
            "kind": "message",
            "ts": "2026-07-15T10:25:00Z",
            "sender": "Alice",
            "body": "Yes, agreed",
            "is_body_empty": False,
            "reply_to": {"sender": "Bob", "preview": "the original question"},
        }
        line = _format_chat_entry(entry)
        assert "↩️" in line
        assert "Bob" in line
        assert "the original question" in line
        assert "Yes, agreed" in line

    def test_message_without_reply_context_no_quote(self):
        from ms365_intent_mcp.formatters import _format_chat_entry
        entry = {
            "kind": "message",
            "ts": "2026-07-15T10:25:00Z",
            "sender": "Alice",
            "body": "hi",
            "is_body_empty": False,
        }
        line = _format_chat_entry(entry)
        assert "↩️" not in line

    def test_call_with_initiator_duration_recording_transcript(self):
        from ms365_intent_mcp.formatters import _format_chat_entry
        entry = {
            "kind": "call",
            "ts": "2026-05-29T10:00:00Z",
            "end_ts": "2026-05-29T10:25:00Z",
            "duration": "25m0s",
            "recording_url": "https://r.example/c.mp4",
            "transcript_ready": True,
            "initiator": "Bawa",
        }
        line = _format_chat_entry(entry)
        assert "Call started by Bawa" in line
        # Bug 4: same-day call shows date once + time range; UTC label on both sides
        assert "2026-05-29 10:00 UTC–10:25 UTC" in line
        assert "25m0s" in line
        assert "[recording](https://r.example/c.mp4)" in line
        assert "transcript ready" in line

    def test_call_no_initiator(self):
        from ms365_intent_mcp.formatters import _format_chat_entry
        entry = {
            "kind": "call",
            "ts": "2026-05-29T10:00:00Z",
            "end_ts": "2026-05-29T10:25:00Z",
            "duration": "25m0s",
            "recording_url": "",
            "transcript_ready": False,
            "initiator": None,
        }
        line = _format_chat_entry(entry)
        assert "**Call**" in line
        assert "started by" not in line

    def test_call_pending_recording_no_link(self):
        from ms365_intent_mcp.formatters import _format_chat_entry
        entry = {
            "kind": "call",
            "ts": "2026-05-29T10:00:00Z",
            "end_ts": "2026-05-29T10:25:00Z",
            "duration": "25m0s",
            "recording_url": "",
            "transcript_ready": False,
            "initiator": "Bawa",
        }
        line = _format_chat_entry(entry)
        assert "[recording]" not in line

    def test_call_single_event_no_end_ts(self):
        from ms365_intent_mcp.formatters import _format_chat_entry
        entry = {
            "kind": "call",
            "ts": "2026-05-29T10:00:00Z",
            "end_ts": "2026-05-29T10:00:00Z",
            "duration": None,
            "recording_url": "",
            "transcript_ready": False,
            "initiator": None,
        }
        line = _format_chat_entry(entry)
        # When start == end, we don't render the dash
        # The format is "({time_range})" — verify no dash inside the parens
        time_range_part = line.split("(", 1)[1].split(")", 1)[0]
        assert "–" not in time_range_part

    def test_call_spanning_dates_renders_both(self):
        """Bug 4: a call that spans midnight (rare) shows both dates."""
        from ms365_intent_mcp.formatters import _format_chat_entry
        entry = {
            "kind": "call",
            "ts": "2026-05-29T23:30:00Z",
            "end_ts": "2026-05-30T00:15:00Z",
            "duration": "45m0s",
            "recording_url": "",
            "transcript_ready": False,
            "initiator": None,
        }
        line = _format_chat_entry(entry)
        # When dates differ, render both end-points fully; UTC label on both
        assert "2026-05-29 23:30 UTC" in line
        assert "2026-05-30 00:15 UTC" in line

    def test_event_member_added(self):
        from ms365_intent_mcp.formatters import _format_chat_entry
        entry = {
            "kind": "event",
            "ts": "2026-05-29T10:00:00Z",
            "event_type": "membersAdded",
            "summary": "Member added: Carol",
        }
        line = _format_chat_entry(entry)
        assert "⚙️" in line
        assert "Member added: Carol" in line
        # Bug 4; UTC label for clarity
        assert "2026-05-29 10:00 UTC" in line

    def test_event_call_unknown(self):
        from ms365_intent_mcp.formatters import _format_chat_entry
        entry = {
            "kind": "event",
            "ts": "2026-05-29T10:00:00Z",
            "event_type": "call_unknown",
            "summary": "Call event (no callId)",
        }
        line = _format_chat_entry(entry)
        assert "Call event (no callId)" in line

    def test_unknown_kind_renders_fallback(self):
        from ms365_intent_mcp.formatters import _format_chat_entry
        line = _format_chat_entry({"kind": "alien"})
        assert "unknown entry: alien" in line

    def test_call_single_event_includes_utc(self):
        from ms365_intent_mcp.formatters import _format_chat_entry
        entry = {
            "kind": "call",
            "ts": "2026-05-29T10:00:00Z",
            "end_ts": "2026-05-29T10:00:00Z",
            "duration": None,
            "recording_url": "",
            "transcript_ready": False,
            "initiator": None,
        }
        line = _format_chat_entry(entry)
        assert "2026-05-29 10:00 UTC" in line

    def test_call_with_null_ts_renders_empty_parens(self):
        from ms365_intent_mcp.formatters import _format_chat_entry
        entry = {
            "kind": "call",
            "ts": "",
            "end_ts": "",
            "duration": None,
            "recording_url": "",
            "transcript_ready": False,
            "initiator": None,
        }
        line = _format_chat_entry(entry)
        # Sanity: function returns a call line, doesn't crash.
        assert "📞" in line
        # No orphan UTC label when there's no timestamp.
        assert "UTC" not in line
        # Empty parens (matching how the message and event branches
        # render null ts via ts_with_tz).
        assert "()" in line

    def test_call_with_drive_metadata_renders_sub_bullets(self):
        """Enriched call entries (drive_id/vroom_url/owner) render as
        sub-bullets beneath the main call line — callers who want a
        ready-to-download URL can grab vroom_url without parsing."""
        from ms365_intent_mcp.formatters import _format_chat_entry
        entry = {
            "kind": "call",
            "ts": "2026-06-30T13:00:00Z",
            "end_ts": "2026-06-30T14:14:00Z",
            "duration": "1h14m",
            "recording_url": "https://sap-my.sharepoint.com/:v:/p/marcus_karlbowski/IQBw",
            "transcript_ready": True,
            "initiator": "Marcus",
            "drive_id": "b!vwRb...",
            "drive_item_id": "01LUWJL4",
            "owner_upn": "marcus_karlbowski",
            "vroom_url": "https://sap-my.sharepoint.com/personal/marcus_karlbowski/_api/v2.0/drives/b!vwRb.../items/01LUWJL4",
        }
        line = _format_chat_entry(entry)
        # Main line still there.
        assert "📞" in line
        assert "[recording]" in line
        # Enriched fields visible as sub-bullets.
        assert "vroom_url:" in line
        assert "drive_id:" in line
        assert "drive_item_id:" in line
        assert "owner:" in line
        assert "marcus_karlbowski" in line
        # Sub-bullets are indented (2 spaces) so they render as a nested list.
        assert "\n  - vroom_url:" in line

    def test_call_without_drive_metadata_stays_single_line(self):
        """Backward-compatible: unenriched call entries render as before."""
        from ms365_intent_mcp.formatters import _format_chat_entry
        entry = {
            "kind": "call",
            "ts": "2026-06-30T13:00:00Z",
            "end_ts": "2026-06-30T14:14:00Z",
            "duration": "1h14m",
            "recording_url": "https://sap-my.sharepoint.com/:v:/p/x/y",
            "transcript_ready": True,
            "initiator": "Marcus",
        }
        line = _format_chat_entry(entry)
        assert "vroom_url" not in line
        assert "drive_id" not in line
        # Still exactly one line.
        assert "\n" not in line


class TestFormatEventTimeRange:
    def test_basic_utc_range(self):
        from ms365_intent_mcp.formatters import _format_event_time_range
        start = {"dateTime": "2026-06-02T07:45:00.0000000", "timeZone": "UTC"}
        end = {"dateTime": "2026-06-02T08:00:00.0000000", "timeZone": "UTC"}
        assert _format_event_time_range(start, end) == "07:45–08:00 UTC"

    def test_named_iana_timezone(self):
        from ms365_intent_mcp.formatters import _format_event_time_range
        start = {"dateTime": "2026-06-02T09:45:00.0000000", "timeZone": "Europe/Berlin"}
        end = {"dateTime": "2026-06-02T10:00:00.0000000", "timeZone": "Europe/Berlin"}
        assert _format_event_time_range(start, end) == "09:45–10:00 Europe/Berlin"

    def test_falls_back_to_end_timezone_if_start_missing(self):
        from ms365_intent_mcp.formatters import _format_event_time_range
        start = {"dateTime": "2026-06-02T07:45:00.0000000"}
        end = {"dateTime": "2026-06-02T08:00:00.0000000", "timeZone": "UTC"}
        assert _format_event_time_range(start, end) == "07:45–08:00 UTC"

    def test_no_timezone_no_suffix(self):
        """Defensive: if Graph somehow omits timeZone, render bare times — no ' None' artifact."""
        from ms365_intent_mcp.formatters import _format_event_time_range
        start = {"dateTime": "2026-06-02T07:45:00.0000000"}
        end = {"dateTime": "2026-06-02T08:00:00.0000000"}
        result = _format_event_time_range(start, end)
        assert result == "07:45–08:00"
        assert "None" not in result

    def test_empty_inputs(self):
        from ms365_intent_mcp.formatters import _format_event_time_range
        assert _format_event_time_range({}, {}) == "–"

    def test_short_datetime_string_passes_through(self):
        """Defensive: if dateTime is shorter than 16 chars, don't crash — render as-is."""
        from ms365_intent_mcp.formatters import _format_event_time_range
        start = {"dateTime": "?", "timeZone": "UTC"}
        end = {"dateTime": "?", "timeZone": "UTC"}
        result = _format_event_time_range(start, end)
        assert "UTC" in result


class TestMentionRegressionAcrossFormatters:
    def test_format_teams_activity_extracts_at_mention(self):
        msgs = [{
            "from": {"user": {"displayName": "Alice"}},
            "body": {"content": '<at id="123">@Avi</at> see this'},
        }]
        result = format_teams_activity_markdown(msgs)
        assert "@Avi see this" in result
        assert "<at" not in result

    def test_format_event_detail_extracts_at_mention(self):
        event = {
            "subject": "Standup",
            "start": {"dateTime": "2026-05-29T09:00:00"},
            "end": {"dateTime": "2026-05-29T09:30:00"},
            "organizer": {"emailAddress": {"name": "Bob"}},
            "location": {"displayName": ""},
            "isOnlineMeeting": False,
            "attendees": [],
            "body": {"content": 'agenda: <at id="123">@Avi</at> updates'},
        }
        result = format_event_detail_markdown(event)
        assert "@Avi" in result

    def test_format_people_extracts_at_mention_in_recent_chat(self):
        people = [{"displayName": "Bawa Kulkarni", "emailAddresses": [{"address": "bawa@example.com"}]}]
        recent_chat = {
            "lastMessagePreview": {
                "body": {"content": '<at id="123">@Avi</at> ping'},
            }
        }
        result = format_people_markdown("bawa", people, [], recent_chat)
        assert "@Avi" in result


class TestFormatOffsetDatetime:
    def test_basic_z_suffix(self):
        from ms365_intent_mcp.formatters import _format_offset_datetime
        assert _format_offset_datetime("2026-05-29T10:00:00Z") == "2026-05-29T10:00 UTC"

    def test_with_milliseconds(self):
        """Graph dateTimeOffset can include ms: '2026-05-29T10:00:00.035Z'."""
        from ms365_intent_mcp.formatters import _format_offset_datetime
        assert _format_offset_datetime("2026-05-29T10:00:00.035Z") == "2026-05-29T10:00 UTC"

    def test_empty_string(self):
        from ms365_intent_mcp.formatters import _format_offset_datetime
        assert _format_offset_datetime("") == ""

    def test_none_input(self):
        """Callers may pass `data.get('createdDateTime')` directly (no default)."""
        from ms365_intent_mcp.formatters import _format_offset_datetime
        assert _format_offset_datetime(None) == ""

    def test_short_string_passes_through(self):
        from ms365_intent_mcp.formatters import _format_offset_datetime
        result = _format_offset_datetime("???")
        assert "UTC" not in result  # don't slap UTC on garbage


class TestFormatEventDatetime:
    def test_basic_utc(self):
        from ms365_intent_mcp.formatters import _format_event_datetime
        dt = {"dateTime": "2026-06-02T07:45:00.0000000", "timeZone": "UTC"}
        assert _format_event_datetime(dt) == "2026-06-02T07:45 UTC"

    def test_named_timezone(self):
        from ms365_intent_mcp.formatters import _format_event_datetime
        dt = {"dateTime": "2026-06-02T07:45:00.0000000", "timeZone": "Europe/Berlin"}
        assert _format_event_datetime(dt) == "2026-06-02T07:45 Europe/Berlin"

    def test_no_timezone_no_suffix(self):
        from ms365_intent_mcp.formatters import _format_event_datetime
        dt = {"dateTime": "2026-06-02T07:45:00.0000000"}
        result = _format_event_datetime(dt)
        assert result == "2026-06-02T07:45"
        assert "None" not in result

    def test_empty_input(self):
        from ms365_intent_mcp.formatters import _format_event_datetime
        assert _format_event_datetime({}) == ""

    def test_missing_datetime_key(self):
        from ms365_intent_mcp.formatters import _format_event_datetime
        dt = {"timeZone": "UTC"}
        assert _format_event_datetime(dt) == " UTC"




def test_utc_pair_becomes_z():
    iso, tz = graph_dt_to_aware_iso(
        {"dateTime": "2026-07-29T14:00:00.0000000", "timeZone": "UTC"}
    )
    assert iso == "2026-07-29T14:00:00+00:00"
    assert tz == "UTC"


class TestEmailAttachmentRendering:
    def test_renders_attachment_lines(self):
        data = {
            "subject": "Bug", "from": {"emailAddress": {"name": "Cust"}},
            "receivedDateTime": "2026-07-20T00:00:00Z",
            "body": {"contentType": "text", "content": "see attached"},
            "_attachments": [
                {"name": "shot.png", "content_type": "image/png", "size": 20480,
                 "is_inline": True, "cid": "a@1", "attachment_id": "i",
                 "kind": "file", "local_path": "/tmp/shot.png", "note": None},
                {"name": "notes.docx", "content_type": "", "size": 0,
                 "is_inline": False, "cid": "", "attachment_id": "j",
                 "kind": "item", "local_path": None,
                 "note": "embedded item — not a downloadable file"},
            ],
        }
        md = format_resolved_content_markdown("email", data)
        assert "shot.png" in md
        assert "/tmp/shot.png" in md
        assert "notes.docx" in md
        assert "not a downloadable file" in md

    def test_no_attachment_section_when_empty(self):
        data = {
            "subject": "plain", "from": {"emailAddress": {"name": "A"}},
            "receivedDateTime": "2026-07-20T00:00:00Z",
            "body": {"contentType": "text", "content": "hi"},
            "_attachments": [],
        }
        md = format_resolved_content_markdown("email", data)
        assert "📎" not in md

    def test_attachments_error_warning_rendered(self):
        data = {
            "subject": "err", "from": {"emailAddress": {"name": "B"}},
            "receivedDateTime": "2026-07-20T00:00:00Z",
            "body": {"contentType": "text", "content": "hi"},
            "_attachments": [],
            "_attachments_error": "403 Forbidden",
        }
        md = format_resolved_content_markdown("email", data)
        assert "403 Forbidden" in md
        assert "⚠️" in md
