"""Tests for meeting composer."""

from unittest.mock import AsyncMock, patch

import pytest

from ms365_intent_mcp.composers.meeting import _resolve_recording_for_event, compose_meeting
from ms365_intent_mcp.permissions import PermissionRegistry


@pytest.fixture
def full_permissions():
    return PermissionRegistry(["Calendars.ReadWrite", "Chat.ReadWrite"])


class TestMeetingById:
    @pytest.mark.asyncio
    async def test_returns_event_details(self, full_permissions):
        client = AsyncMock()
        client.get = AsyncMock(return_value=_full_event())

        _, result = await compose_meeting(
            client=client,
            permissions=full_permissions,
            identifier="event-id-123",
            timezone="Europe/Berlin",
        )
        assert "Team Sync" in result
        assert "Alice" in result


class TestMeetingBySubject:
    @pytest.mark.asyncio
    async def test_searches_by_subject(self, full_permissions):
        client = AsyncMock()

        async def _get(endpoint, params=None, headers=None):
            if "calendarView" in endpoint:
                return {"value": [_full_event()]}
            return _full_event()

        client.get = AsyncMock(side_effect=_get)
        _, result = await compose_meeting(
            client=client,
            permissions=full_permissions,
            identifier="Team Sync",
            timezone="Europe/Berlin",
        )
        assert "Team Sync" in result


class TestMeetingNext:
    @pytest.mark.asyncio
    async def test_next_finds_upcoming(self, full_permissions):
        client = AsyncMock()
        client.get = AsyncMock(return_value={"value": [_full_event()]})

        _, result = await compose_meeting(
            client=client,
            permissions=full_permissions,
            identifier="next",
            timezone="Europe/Berlin",
        )
        assert "Team Sync" in result


def _full_event():
    return {
        "id": "event-id-123",
        "subject": "Team Sync",
        "start": {"dateTime": "2026-05-15T14:00:00"},
        "end": {"dateTime": "2026-05-15T14:30:00"},
        "location": {"displayName": "Room B"},
        "isOnlineMeeting": True,
        "onlineMeeting": {"joinUrl": "https://teams.microsoft.com/l/meetup-join/123"},
        "organizer": {"emailAddress": {"name": "Alice", "address": "alice@example.com"}},
        "attendees": [
            {
                "emailAddress": {"name": "Bob", "address": "bob@example.com"},
                "status": {"response": "accepted"},
            },
        ],
        "body": {"content": "<p>Agenda: review Q2 progress</p>", "contentType": "html"},
    }


class TestResolveRecordingForEvent:
    @pytest.mark.asyncio
    async def test_returns_none_when_not_online(self):
        client = AsyncMock()
        event = {"isOnlineMeeting": False}
        result = await _resolve_recording_for_event(client, event)
        assert result is None
        client.get.assert_not_called()

    @pytest.mark.asyncio
    async def test_returns_none_when_join_url_has_no_thread(self):
        """joinUrl without a `meetup-join/19:...@thread.v2` segment can't be
        resolved to a chat — return None, don't call Graph."""
        client = AsyncMock()
        event = {
            "isOnlineMeeting": True,
            "onlineMeeting": {"joinUrl": "https://teams.microsoft.com/l/meetup-join/nope"},
        }
        result = await _resolve_recording_for_event(client, event)
        assert result is None
        client.get.assert_not_called()

    @pytest.mark.asyncio
    async def test_extracts_thread_from_join_url_and_calls_chat(self):
        """The joinUrl carries the same threadId that chatInfo would give us."""
        client = AsyncMock()
        client.get = AsyncMock(
            return_value={
                "value": [
                    {
                        "createdDateTime": "2026-06-30T13:00:00Z",
                        "eventDetail": {
                            "@odata.type": "#microsoft.graph.callRecordingEventMessageDetail",
                            "callRecordingStatus": "success",
                            "callRecordingUrl": "https://sap-my.sharepoint.com/:v:/p/marcus_karlbowski/IQBw",
                            "callRecordingDisplayName": "[NOVA] Refinement.mp4",
                        },
                    },
                ]
            }
        )
        # Guard the enrichment step — we're testing thread extraction here.
        with patch("ms365_intent_mcp.composers.meeting._enrich_call_recording") as enrich:
            enrich.return_value = None
            event = {
                "isOnlineMeeting": True,
                "onlineMeeting": {
                    "joinUrl": (
                        "https://teams.microsoft.com/l/meetup-join/"
                        "19%3Ameeting_abc%40thread.v2/0?context=..."
                    ),
                },
            }
            result = await _resolve_recording_for_event(client, event)

        assert result is not None
        assert (
            result["recording_url"] == "https://sap-my.sharepoint.com/:v:/p/marcus_karlbowski/IQBw"
        )
        assert result["display_name"] == "[NOVA] Refinement.mp4"
        # Chat was fetched using the thread id extracted from the joinUrl.
        endpoint = client.get.call_args[0][0]
        assert endpoint == "/chats/19:meeting_abc@thread.v2/messages"

    @pytest.mark.asyncio
    async def test_returns_none_when_chat_has_no_recording_events(self):
        client = AsyncMock()
        client.get = AsyncMock(return_value={"value": []})
        event = {
            "isOnlineMeeting": True,
            "onlineMeeting": {
                "joinUrl": (
                    "https://teams.microsoft.com/l/meetup-join/19%3Ameeting_abc%40thread.v2/0"
                ),
            },
        }
        result = await _resolve_recording_for_event(client, event)
        assert result is None

    @pytest.mark.asyncio
    async def test_transcript_ready_when_transcript_event_present(self):
        client = AsyncMock()
        client.get = AsyncMock(
            return_value={
                "value": [
                    {
                        "createdDateTime": "2026-06-30T13:00:00Z",
                        "eventDetail": {
                            "@odata.type": "#microsoft.graph.callRecordingEventMessageDetail",
                            "callRecordingStatus": "success",
                            "callRecordingUrl": "https://sap-my.sharepoint.com/:v:/p/x/y",
                        },
                    },
                    {
                        "createdDateTime": "2026-06-30T13:05:00Z",
                        "eventDetail": {
                            "@odata.type": "#microsoft.graph.callTranscriptEventMessageDetail",
                        },
                    },
                ]
            }
        )
        with patch("ms365_intent_mcp.composers.meeting._enrich_call_recording") as enrich:
            enrich.return_value = None
            event = {
                "isOnlineMeeting": True,
                "onlineMeeting": {
                    "joinUrl": "https://teams.microsoft.com/l/meetup-join/19%3Ameeting_abc%40thread.v2/0",
                },
            }
            result = await _resolve_recording_for_event(client, event)

        assert result is not None
        assert result["transcript_ready"] is True

    @pytest.mark.asyncio
    async def test_prefers_latest_success_recording_url(self):
        """When multiple callRecording events fire, the freshest success wins."""
        client = AsyncMock()
        client.get = AsyncMock(
            return_value={
                "value": [
                    {
                        "createdDateTime": "2026-06-30T13:00:00Z",
                        "eventDetail": {
                            "@odata.type": "#microsoft.graph.callRecordingEventMessageDetail",
                            "callRecordingStatus": "initial",
                            "callRecordingUrl": "https://old",
                        },
                    },
                    {
                        "createdDateTime": "2026-06-30T13:05:00Z",
                        "eventDetail": {
                            "@odata.type": "#microsoft.graph.callRecordingEventMessageDetail",
                            "callRecordingStatus": "success",
                            "callRecordingUrl": "https://new",
                        },
                    },
                ]
            }
        )
        with patch("ms365_intent_mcp.composers.meeting._enrich_call_recording"):
            event = {
                "isOnlineMeeting": True,
                "onlineMeeting": {
                    "joinUrl": "https://teams.microsoft.com/l/meetup-join/19%3Ameeting_abc%40thread.v2/0",
                },
            }
            result = await _resolve_recording_for_event(client, event)

        assert result is not None
        assert result["recording_url"] == "https://new"


class TestMeetingDetailTimezones:
    @pytest.mark.asyncio
    async def test_meeting_detail_start_is_offset_aware(self, full_permissions):
        """UTC event: start/end must be offset-aware ISO and carry _timezone siblings."""
        event = {
            "id": "utc-event",
            "subject": "UTC Meeting",
            "start": {"dateTime": "2026-07-29T14:00:00.0000000", "timeZone": "UTC"},
            "end": {"dateTime": "2026-07-29T14:30:00.0000000", "timeZone": "UTC"},
            "location": {"displayName": ""},
            "isOnlineMeeting": False,
            "organizer": {"emailAddress": {"name": "Avi", "address": "avi@example.com"}},
            "attendees": [],
            "body": {"content": "", "contentType": "html"},
        }
        client = AsyncMock()
        client.get = AsyncMock(return_value=event)
        data, _ = await compose_meeting(
            client=client,
            permissions=full_permissions,
            identifier="utc-event",
            timezone="UTC",
        )
        assert data["start"] == "2026-07-29T14:00:00+00:00"
        assert data["end"] == "2026-07-29T14:30:00+00:00"
        assert data["start_timezone"] == "UTC"
        assert data["end_timezone"] == "UTC"


class TestMeetingWithRecording:
    @pytest.mark.asyncio
    async def test_meeting_output_includes_recording_when_available(self, full_permissions):
        """End-to-end at the composer level: recording metadata reaches
        the rendered event detail."""
        event = _full_event()
        # Overwrite joinUrl with one that has an extractable thread id.
        event["onlineMeeting"]["joinUrl"] = (
            "https://teams.microsoft.com/l/meetup-join/19%3Ameeting_xyz%40thread.v2/0"
        )
        client = AsyncMock()

        async def _get(endpoint, params=None, headers=None):
            if "/messages" in endpoint:
                return {
                    "value": [
                        {
                            "createdDateTime": "2026-05-15T14:30:00Z",
                            "eventDetail": {
                                "@odata.type": "#microsoft.graph.callRecordingEventMessageDetail",
                                "callRecordingStatus": "success",
                                "callRecordingUrl": "https://sap-my.sharepoint.com/:v:/p/alice/IQBw",
                                "callRecordingDisplayName": "Team Sync-recording.mp4",
                            },
                        },
                    ]
                }
            if endpoint.startswith("/shares/"):
                return {
                    "id": "01ABC",
                    "parentReference": {"driveId": "b!xxx"},
                }
            return event

        client.get = AsyncMock(side_effect=_get)

        _, result = await compose_meeting(
            client=client,
            permissions=full_permissions,
            identifier="event-id-123",
            timezone="Europe/Berlin",
        )

        assert "Recording" in result
        assert "https://sap-my.sharepoint.com/:v:/p/alice/IQBw" in result
        assert "Team Sync-recording.mp4" in result
        assert "b!xxx" in result
        assert "01ABC" in result
        assert "alice" in result

    @pytest.mark.asyncio
    async def test_meeting_output_omits_recording_block_when_absent(self, full_permissions):
        """No recording events → no `Recording:` block in the output.
        Previous behavior preserved for meetings without recordings."""
        event = _full_event()
        event["onlineMeeting"]["joinUrl"] = (
            "https://teams.microsoft.com/l/meetup-join/19%3Ameeting_xyz%40thread.v2/0"
        )
        client = AsyncMock()

        async def _get(endpoint, params=None, headers=None):
            if "/messages" in endpoint:
                return {"value": []}
            return event

        client.get = AsyncMock(side_effect=_get)
        _, result = await compose_meeting(
            client=client,
            permissions=full_permissions,
            identifier="event-id-123",
            timezone="Europe/Berlin",
        )
        assert "Recording" not in result
