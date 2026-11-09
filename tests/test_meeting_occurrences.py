"""Synthetic recurrence regressions: no live IDs, content, or Graph access."""

from datetime import UTC, date, datetime
from unittest.mock import AsyncMock, MagicMock, patch

import pytest
from pydantic import ValidationError

from ms365_intent_mcp.composers.meeting import compose_meeting
from ms365_intent_mcp.graph import GraphAPIError, GraphClient
from ms365_intent_mcp.intent._helpers import IntentError
from ms365_intent_mcp.intent._shared import ErrorResponse
from ms365_intent_mcp.intent.meeting.impl import _meeting_impl
from ms365_intent_mcp.intent.meeting.schemas import MeetingPayload
from ms365_intent_mcp.intent.my_day.impl import _my_day_impl
from ms365_intent_mcp.intent.my_day.schemas import EventSummary, MyDayPayload
from ms365_intent_mcp.permissions import PermissionRegistry


def occurrence(day, *, timezone="UTC", time="09:00:00", **extra):
    return {
        "id": f"synthetic-occurrence-{day}",
        "type": "occurrence",
        "subject": "Synthetic Weekly Sync",
        "start": {"dateTime": f"{day}T{time}", "timeZone": timezone},
        "end": {"dateTime": f"{day}T10:00:00", "timeZone": timezone},
        "organizer": {"emailAddress": {"name": "Synthetic Organizer"}},
        "attendees": [],
        **extra,
    }


def calendar_client(events):
    client = AsyncMock(spec=GraphClient)

    async def get(endpoint, params=None, headers=None):
        if endpoint.startswith("/me/events/"):
            matches = [e for e in events if e["id"] == endpoint.removeprefix("/me/events/")]
            if matches:
                return matches[0]
            raise GraphAPIError(404, "ErrorItemNotFound", "Synthetic ID miss")
        if endpoint == "/me/calendarView":
            # Intentionally return neighbors too: callers must validate the actual start date.
            return {"value": events}
        return {"value": []}

    async def get_all(*args, **kwargs):
        return await GraphClient.get_all(client, *args, **kwargs)

    client.get.side_effect = get
    client.get_all.side_effect = get_all
    return client


def context(client, timezone="UTC"):
    ctx = MagicMock()
    ctx.request_context.lifespan_context = {
        "config": MagicMock(default_timezone=timezone),
        "client": client,
        "permissions": PermissionRegistry(["Calendars.ReadWrite"]),
    }
    return ctx


@pytest.fixture
def recurrences():
    return [occurrence(f"2026-11-{day}") for day in ("02", "09", "16")]


@pytest.mark.asyncio
async def test_omitted_date_reproduces_earliest_recurrence_and_preserves_compatibility(recurrences):
    client = calendar_client(recurrences)
    with patch("ms365_intent_mcp.composers.meeting.datetime") as clock:
        clock.now.return_value = datetime(2026, 11, 9, 12)
        data, _ = await compose_meeting(client, None, "Synthetic Weekly Sync", "UTC")
    assert data["id"] == "synthetic-occurrence-2026-11-02"
    params = client.get.call_args.kwargs["params"]
    assert params["startDateTime"] == "2026-11-02T00:00:00"
    assert params["endDateTime"] == "2026-11-16T23:59:59"


@pytest.mark.asyncio
@pytest.mark.parametrize("day", ["02", "09", "16"])
async def test_explicit_date_selects_only_correct_recurrence(recurrences, day):
    client = calendar_client(recurrences)
    requested = f"2026-11-{day}"
    response = await _meeting_impl(
        context(client), MeetingPayload(identifier="Synthetic Weekly Sync", date=requested)
    )
    assert response.id == f"synthetic-occurrence-{requested}"
    assert response.start.date().isoformat() == requested
    params = client.get_all.call_args.kwargs["params"]
    assert params["startDateTime"] == f"{requested}T00:00:00+00:00"
    assert client.get_all.call_args.kwargs["headers"] == {"Prefer": 'outlook.timezone="UTC"'}


@pytest.mark.asyncio
async def test_date_miss_returns_not_found_without_recording_or_fabricated_meeting(recurrences):
    client = calendar_client(recurrences)
    response = await _meeting_impl(
        context(client), MeetingPayload(identifier="Synthetic Weekly Sync", date="2026-11-10")
    )
    assert isinstance(response, ErrorResponse)
    assert response.code == "not_found"
    assert "2026-11-10" in response.message
    assert all("/chats/" not in c.args[0] for c in client.get.call_args_list)


@pytest.mark.parametrize(
    "value",
    [
        "20261109",
        "2026-11-9",
        "2026-02-30",
        "2026-11-09T00:00:00",
        "2026-W46-1",
        1794182400,
        datetime(2026, 11, 9, tzinfo=UTC),
    ],
)
def test_malformed_payload_dates_rejected(value):
    with pytest.raises(ValidationError):
        MeetingPayload(identifier="Synthetic Weekly Sync", date=value)


def test_date_objects_and_omission_supported():
    assert MeetingPayload(identifier="x", date=date(2026, 11, 9)).date == date(2026, 11, 9)
    assert MeetingPayload(identifier="x").date is None
    with pytest.raises(ValidationError):
        MeetingPayload(identifier="NEXT", date="2026-11-09")


@pytest.mark.asyncio
@pytest.mark.parametrize("value", ["20261109", "2026-02-30", "2026-11-9"])
async def test_composer_also_validates_dates_before_graph(value):
    client = calendar_client([])
    with pytest.raises(IntentError, match="Invalid occurrence date"):
        await compose_meeting(client, None, "Synthetic Weekly Sync", "UTC", date=value)
    client.get.assert_not_called()


@pytest.mark.asyncio
async def test_direct_id_omitted_date_and_matching_date_supported(recurrences):
    client = calendar_client(recurrences)
    identifier = recurrences[1]["id"]
    for requested in (None, "2026-11-09"):
        data, _ = await compose_meeting(client, None, identifier, "UTC", date=requested)
        assert data["id"] == identifier
    client.get_all.assert_not_called()


@pytest.mark.asyncio
async def test_direct_id_date_mismatch_cannot_fall_through_to_subject(recurrences):
    client = calendar_client(recurrences)
    with pytest.raises(IntentError) as exc:
        await compose_meeting(client, None, recurrences[0]["id"], "UTC", date="2026-11-09")
    assert exc.value.code == "not_found"
    client.get_all.assert_not_called()


@pytest.mark.asyncio
async def test_series_master_id_with_date_rejected():
    event = occurrence("2026-11-09", type="seriesMaster")
    client = calendar_client([event])
    response = await _meeting_impl(
        context(client), MeetingPayload(identifier=event["id"], date="2026-11-09")
    )
    assert response.code == "validation_error"
    client.get_all.assert_not_called()


@pytest.mark.asyncio
@pytest.mark.parametrize("include_bodies", [False, True])
async def test_typed_my_day_ids_roundtrip_to_meeting_without_rendering_ids(include_bodies):
    event = occurrence("2026-11-09", body={"content": "Synthetic agenda"})
    client = calendar_client([event])
    ctx = context(client)
    day = await _my_day_impl(ctx, MyDayPayload(date="2026-11-09", include_bodies=include_bodies))
    assert day.events[0].id == event["id"]
    assert event["id"] not in day.rendered_markdown
    assert "id" in client.get.call_args.kwargs["params"]["$select"].split(",")
    meeting = await _meeting_impl(ctx, MeetingPayload(identifier=day.events[0].id))
    assert meeting.id == event["id"]
    assert meeting.start == day.events[0].start


def test_event_summary_without_id_remains_valid():
    event = EventSummary(subject="Synthetic", start="2026-11-09", end="2026-11-09")
    assert event.id == ""


@pytest.mark.asyncio
async def test_filter_fallback_and_pagination_still_check_date(recurrences):
    client = calendar_client(recurrences)
    client.get.side_effect = [
        GraphAPIError(400, "ErrorInvalidUrlQueryFilter", "Unsupported filter"),
        {"value": [recurrences[0]], "@odata.nextLink": "https://graph.microsoft.com/v1.0/next"},
        {"value": [recurrences[1], recurrences[2]]},
    ]
    data, _ = await compose_meeting(client, None, "Synthetic Weekly Sync", "UTC", "2026-11-09")
    assert data["id"] == recurrences[1]["id"]
    assert "$filter" not in client.get_all.call_args.kwargs["params"]
    assert client.get.call_args.args[0] == "https://graph.microsoft.com/v1.0/next"


@pytest.mark.asyncio
@pytest.mark.parametrize("status", [403, 429, 503])
async def test_explicit_lookup_does_not_mask_graph_errors(status):
    client = calendar_client([])
    client.get.side_effect = GraphAPIError(status, "SyntheticError", "Synthetic failure")
    response = await _meeting_impl(
        context(client), MeetingPayload(identifier="Synthetic Weekly Sync", date="2026-11-09")
    )
    assert response.code == ("rate_limited" if status in (429, 503) else "graph_api_error")
    assert client.get.call_count == 1


@pytest.mark.asyncio
@pytest.mark.parametrize("timezone", ["Europe/Berlin", "W. Europe Standard Time"])
async def test_local_midnight_and_dst_bounds(timezone):
    # Oct 25 is the 25-hour DST fall-back day. UTC Oct24 22:30 is local Oct25 00:30.
    correct = occurrence("2026-10-24", time="22:30:00", id="synthetic-local-midnight")
    previous = occurrence("2026-10-24", time="21:30:00", id="synthetic-previous-day")
    following = occurrence("2026-10-25", time="23:00:00")
    client = calendar_client([previous, correct, following])
    data, _ = await compose_meeting(client, None, "Synthetic Weekly Sync", timezone, "2026-10-25")
    assert data["id"] == correct["id"]
    params = client.get_all.call_args.kwargs["params"]
    assert params["startDateTime"] == "2026-10-25T00:00:00+02:00"
    assert params["endDateTime"] == "2026-10-26T00:00:00+01:00"


@pytest.mark.asyncio
@pytest.mark.parametrize("recording_day,expected", [("09", True), ("02", False)])
async def test_recording_guard_compares_actual_selected_occurrence(recording_day, expected):
    event = occurrence(
        "2026-11-09",
        isOnlineMeeting=True,
        onlineMeeting={
            "joinUrl": "https://teams.microsoft.com/l/meetup-join/19:synthetic@thread.v2/0"
        },
    )
    client = calendar_client([occurrence("2026-11-02"), event, occurrence("2026-11-16")])
    original_get = client.get.side_effect

    async def get(endpoint, params=None, headers=None):
        if endpoint.startswith("/chats/"):
            return {
                "value": [
                    {
                        "createdDateTime": f"2026-11-{recording_day}T10:00:00Z",
                        "eventDetail": {
                            "@odata.type": "#microsoft.graph.callRecordingEventMessageDetail",
                            "callRecordingStatus": "success",
                            "callRecordingUrl": "https://example.sharepoint.com/synthetic-recording",
                            "callRecordingDisplayName": (
                                f"Synthetic Weekly Sync-202611{recording_day}"
                                "_090000-Meeting Recording.mp4"
                            ),
                            "callId": "synthetic-call",
                        },
                    }
                ]
            }
        return await original_get(endpoint, params, headers)

    client.get.side_effect = get
    with patch("ms365_intent_mcp.composers.meeting._enrich_call_recording", new=AsyncMock()):
        response = await _meeting_impl(
            context(client), MeetingPayload(identifier="Synthetic Weekly Sync", date="2026-11-09")
        )
    assert response.id == event["id"]
    assert response.recording.occurrence_date == "2026-11-09"
    assert response.recording.recording_date == f"2026-11-{recording_day}"
    assert response.recording.date_matches_occurrence is expected
    if not expected:
        assert "**stale:**" in response.rendered_markdown


@pytest.mark.asyncio
async def test_single_word_subject_falls_back_from_id_miss():
    event = occurrence("2026-11-09", subject="Synthetic")
    client = calendar_client([event])
    data, _ = await compose_meeting(client, None, "Synthetic", "UTC", "2026-11-09")
    assert data["id"] == event["id"]
    assert client.get.call_args_list[0].args[0] == "/me/events/Synthetic"


@pytest.mark.asyncio
async def test_explicit_date_nonrecurring_and_all_day_events():
    for event in (
        occurrence("2026-11-09", type="singleInstance"),
        occurrence(
            "2026-11-09",
            type="singleInstance",
            isAllDay=True,
            start={"dateTime": "2026-11-09", "timeZone": "UTC"},
        ),
    ):
        client = calendar_client([event])
        data, _ = await compose_meeting(client, None, "Synthetic Weekly Sync", "UTC", "2026-11-09")
        assert data["id"] == event["id"]


@pytest.mark.asyncio
async def test_id_path_is_encoded_without_reinterpreting_it_as_subject():
    event = occurrence("2026-11-09", id="synthetic/id+value=")
    client = calendar_client([])
    client.get.return_value = event
    client.get.side_effect = None
    data, _ = await compose_meeting(client, None, event["id"], "UTC", "2026-11-09")
    assert data["id"] == event["id"]
    assert client.get.call_args.args[0] == "/me/events/synthetic%2Fid%2Bvalue%3D"
    client.get_all.assert_not_called()


@pytest.mark.asyncio
async def test_truncated_calendar_is_not_a_definitive_date_miss():
    client = calendar_client([])
    client.get_all.side_effect = None
    client.get_all.return_value = ([], True)
    response = await _meeting_impl(
        context(client), MeetingPayload(identifier="Synthetic Weekly Sync", date="2026-11-09")
    )
    assert response.code == "graph_api_error"
    assert "truncated" in response.message


@pytest.mark.asyncio
async def test_unknown_timezone_rejected_before_graph():
    client = calendar_client([])
    response = await _meeting_impl(
        context(client, "Synthetic/Unknown"),
        MeetingPayload(identifier="Synthetic Weekly Sync", date="2026-11-09"),
    )
    assert response.code == "validation_error"
    client.get.assert_not_called()


@pytest.mark.asyncio
async def test_negative_offset_day_boundary():
    previous = occurrence("2026-11-09", time="07:59:59", id="synthetic-previous")
    correct = occurrence("2026-11-09", time="08:00:00", id="synthetic-midnight")
    following = occurrence("2026-11-10", time="08:00:00", id="synthetic-following")
    client = calendar_client([previous, correct, following])
    data, _ = await compose_meeting(
        client, None, "Synthetic Weekly Sync", "America/Los_Angeles", "2026-11-09"
    )
    assert data["id"] == correct["id"]
    params = client.get_all.call_args.kwargs["params"]
    assert params["startDateTime"] == "2026-11-09T00:00:00-08:00"
    assert params["endDateTime"] == "2026-11-10T00:00:00-08:00"
