"""Tests for markdown formatters."""

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
)


class TestFormatEventsMarkdown:
    def test_empty_list(self):
        result = format_events_markdown([])
        assert "No events" in result

    def test_single_event(self):
        events = [{
            "subject": "Standup",
            "start": {"dateTime": "2026-05-15T09:00:00"},
            "end": {"dateTime": "2026-05-15T09:30:00"},
            "location": {"displayName": "Room A"},
            "isOnlineMeeting": True,
            "attendees": [{"emailAddress": {"name": "Bob"}}],
            "organizer": {"emailAddress": {"name": "Alice"}},
        }]
        result = format_events_markdown(events)
        assert "Standup" in result
        assert "09:00" in result
        assert "Room A" in result

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
            "start": {"dateTime": "2026-05-15T14:00:00"},
            "end": {"dateTime": "2026-05-15T14:30:00"},
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
            "start": {"dateTime": "2026-05-16T10:00:00"},
            "end": {"dateTime": "2026-05-16T10:30:00"},
            "isOnlineMeeting": False,
            "onlineMeeting": None,
        }
        result = format_event_created_markdown(event)
        assert "New Meeting" in result
        assert "2026-05-16" in result


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
        people = [{"displayName": "Alice Smith", "emailAddresses": [{"address": "alice@sap.com"}]}]
        emails = [{"subject": "Project update", "from": {"emailAddress": {"name": "Bob"}}, "receivedDateTime": "2026-05-15T10:00:00Z"}]
        result = format_people_markdown("alice", people, emails, None)
        assert "Alice Smith" in result
        assert "Project update" in result

    def test_with_teams_chat(self):
        people = [{"displayName": "Alice Smith", "emailAddresses": [{"address": "alice@sap.com"}]}]
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
                    "from": {"emailAddress": {"name": "Finance", "address": "finance@sap.com"}},
                    "receivedDateTime": "2026-05-10T09:00:00Z",
                    "bodyPreview": "Please find attached the Q2 budget...",
                }
            }
        ]
        result = format_search_results_markdown("budget", hits)
        assert "Q2 Budget Review" in result
        assert "Finance" in result


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
        assert "10:00" in result
        assert "100" in result


class TestFormatResolvedContentMarkdown:
    def test_email_type(self):
        data = {
            "subject": "Hello",
            "from": {"emailAddress": {"name": "Bob", "address": "bob@sap.com"}},
            "receivedDateTime": "2026-05-15T08:00:00Z",
            "bodyPreview": "Hi there",
        }
        result = format_resolved_content_markdown("email", data)
        assert "Hello" in result
        assert "Bob" in result

    def test_sharepoint_page_type(self):
        data = {"displayName": "Project Overview", "webUrl": "https://sap.sharepoint.com/sites/proj"}
        result = format_resolved_content_markdown("sharepoint_page", data)
        assert "Project Overview" in result

    def test_onedrive_file_type(self):
        data = {"name": "report.xlsx", "size": 20480, "webUrl": "https://sap-my.sharepoint.com/files/1"}
        result = format_resolved_content_markdown("onedrive_file", data)
        assert "report.xlsx" in result


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
            "messages": [
                {"from": {"user": {"displayName": "Alice"}},
                 "body": {"content": "Hello world"},
                 "createdDateTime": "2026-05-26T10:05:00Z"},
                {"from": {"user": {"displayName": "Bob"}},
                 "body": {"content": "Hi"},
                 "createdDateTime": "2026-05-26T10:00:00Z"},
            ],
            "meeting": {
                "subject": "Project Sync",
                "start": {"dateTime": "2026-05-26T10:00:00"},
                "end": {"dateTime": "2026-05-26T10:30:00"},
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
        long_body = "x" * 600
        data = self._data(messages=[
            {"from": {"user": {"displayName": "Alice"}},
             "body": {"content": long_body},
             "createdDateTime": "2026-05-26T10:00:00Z"},
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
        data = self._data(messages=[], _messages_error="Microsoft service error")
        result = format_resolved_content_markdown("chat_thread", data)
        assert "⚠️" in result
        assert "Microsoft service error" in result

    def test_empty_message_body_renders_placeholder(self):
        data = self._data(messages=[
            {"from": {"user": {"displayName": "Alice"}},
             "body": {"content": ""},
             "createdDateTime": "2026-05-26T10:00:00Z"},
        ])
        result = format_resolved_content_markdown("chat_thread", data)
        assert "Alice" in result
        assert "(no content)" in result

    def test_html_stripped_from_message_body(self):
        data = self._data(messages=[
            {"from": {"user": {"displayName": "Alice"}},
             "body": {"content": "<p>Hello <b>world</b></p>"},
             "createdDateTime": "2026-05-26T10:00:00Z"},
        ])
        result = format_resolved_content_markdown("chat_thread", data)
        assert "Hello world" in result
        assert "<p>" not in result
        assert "<b>" not in result
