"""Golden JSON Schema snapshots for compose payloads and response models.

8 snapshot tests — one per schema. First run writes the JSON file;
subsequent runs assert byte-equality.  Use ``SNAPSHOT_UPDATE=1`` to
refresh a snapshot after an intentional schema change.

Gate confirmation: ``compose_payload.json`` must contain both ``oneOf``
and ``discriminator`` to confirm that FastMCP/Pydantic renders the
discriminated union correctly (Task 1 gate, re-verified here).
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from pydantic import TypeAdapter

from ms365_intent_mcp.intent._shared import ErrorResponse
from ms365_intent_mcp.intent.compose.schemas import (
    ComposeEmail,
    ComposeEvent,
    ComposePayload,
    ComposeTeamsMessage,
    EmailDraftCreated,
    EventCreated,
    TeamsMessageSent,
)
from ms365_intent_mcp.intent.find.schemas import (
    EmailHit,
    FileHit,
    FindPayload,
    FindResults,
    MessageHit,
    PageHit,
    SearchHit,
)
from ms365_intent_mcp.intent.meeting.schemas import (
    AttendeeStatus,
    MeetingDetail,
    MeetingPayload,
    OnlineMeetingInfo,
    PersonRef,
    RecordingMetadata,
)
from ms365_intent_mcp.intent.my_day.schemas import (
    EventSummary,
    MailSummary,
    MyDayPayload,
    MyDaySummary,
    TeamsActivitySummary,
)
from ms365_intent_mcp.intent.people.schemas import (
    ChatPreview,
    MailPreview,
    PeoplePayload,
    PersonDetail,
)
from ms365_intent_mcp.intent.resolve.schemas import (
    ChannelMessageContent,
    ChatMessageContent,
    ChatThreadContent,
    EmailContent,
    MeetingContent,
    OneDriveFileContent,
    ResolvedContent,
    ResolvedContentData,
    ResolvePayload,
    SharePointPageContent,
)
from ms365_intent_mcp.intent.schedule.schemas import (
    ScheduleConstraints,
    SchedulePayload,
    ScheduleSuggestions,
    TimeSlot,
)
from ms365_intent_mcp.intent.whats_new.schemas import (
    EventSummary as WhatsNewEventSummary,
)
from ms365_intent_mcp.intent.whats_new.schemas import (
    MailItem,
    TeamsItem,
    WhatsNewPayload,
    WhatsNewSummary,
)

_SNAPSHOT_DIR = Path(__file__).parent / "snapshots" / "schemas"


# ---------------------------------------------------------------------------
# Payload union
# ---------------------------------------------------------------------------


def test_compose_payload_snapshot(snapshot: Any) -> None:
    schema = TypeAdapter(ComposePayload).json_schema()
    snapshot("compose_payload", schema)

    # Gate: discriminated union must render with both keywords
    on_disk = json.loads((_SNAPSHOT_DIR / "compose_payload.json").read_text())
    assert "oneOf" in on_disk, "compose_payload.json must contain 'oneOf'"
    assert "discriminator" in on_disk, "compose_payload.json must contain 'discriminator'"


# ---------------------------------------------------------------------------
# Payload variants
# ---------------------------------------------------------------------------


def test_compose_email_variant(snapshot: Any) -> None:
    snapshot("compose_email", ComposeEmail.model_json_schema())


def test_compose_event_variant(snapshot: Any) -> None:
    snapshot("compose_event", ComposeEvent.model_json_schema())


def test_compose_teams_message_variant(snapshot: Any) -> None:
    snapshot("compose_teams_message", ComposeTeamsMessage.model_json_schema())


# ---------------------------------------------------------------------------
# Response models
# ---------------------------------------------------------------------------


def test_email_draft_created(snapshot: Any) -> None:
    snapshot("email_draft_created", EmailDraftCreated.model_json_schema())


def test_event_created(snapshot: Any) -> None:
    snapshot("event_created", EventCreated.model_json_schema())


def test_teams_message_sent(snapshot: Any) -> None:
    snapshot("teams_message_sent", TeamsMessageSent.model_json_schema())


def test_error_response(snapshot: Any) -> None:
    snapshot("error_response", ErrorResponse.model_json_schema())


# ---------------------------------------------------------------------------
# my_day schemas (Task 3)
# ---------------------------------------------------------------------------


def test_my_day_payload(snapshot: Any) -> None:
    snapshot("my_day_payload", MyDayPayload.model_json_schema())


def test_my_day_summary(snapshot: Any) -> None:
    snapshot("my_day_summary", MyDaySummary.model_json_schema())


def test_event_summary(snapshot: Any) -> None:
    snapshot("event_summary", EventSummary.model_json_schema())


def test_mail_summary(snapshot: Any) -> None:
    snapshot("mail_summary", MailSummary.model_json_schema())


def test_teams_activity_summary(snapshot: Any) -> None:
    snapshot("teams_activity_summary", TeamsActivitySummary.model_json_schema())


# ---------------------------------------------------------------------------
# people schemas (Task 5)
# ---------------------------------------------------------------------------


def test_people_payload(snapshot: Any) -> None:
    snapshot("people_payload", PeoplePayload.model_json_schema())


def test_person_detail(snapshot: Any) -> None:
    snapshot("person_detail", PersonDetail.model_json_schema())


def test_mail_preview(snapshot: Any) -> None:
    snapshot("mail_preview", MailPreview.model_json_schema())


def test_chat_preview(snapshot: Any) -> None:
    snapshot("chat_preview", ChatPreview.model_json_schema())


# ---------------------------------------------------------------------------
# schedule schemas (Task 6)
# ---------------------------------------------------------------------------


def test_schedule_payload(snapshot: Any) -> None:
    snapshot("schedule_payload", SchedulePayload.model_json_schema())


def test_schedule_suggestions(snapshot: Any) -> None:
    snapshot("schedule_suggestions", ScheduleSuggestions.model_json_schema())


def test_time_slot(snapshot: Any) -> None:
    snapshot("time_slot", TimeSlot.model_json_schema())


def test_schedule_constraints(snapshot: Any) -> None:
    snapshot("schedule_constraints", ScheduleConstraints.model_json_schema())


# ---------------------------------------------------------------------------
# meeting schemas (Task 4)
# ---------------------------------------------------------------------------


def test_meeting_payload(snapshot: Any) -> None:
    snapshot("meeting_payload", MeetingPayload.model_json_schema())


def test_meeting_detail(snapshot: Any) -> None:
    snapshot("meeting_detail", MeetingDetail.model_json_schema())


def test_person_ref(snapshot: Any) -> None:
    snapshot("person_ref", PersonRef.model_json_schema())


def test_attendee_status(snapshot: Any) -> None:
    snapshot("attendee_status", AttendeeStatus.model_json_schema())


def test_online_meeting_info(snapshot: Any) -> None:
    snapshot("online_meeting_info", OnlineMeetingInfo.model_json_schema())


def test_recording_metadata(snapshot: Any) -> None:
    snapshot("recording_metadata", RecordingMetadata.model_json_schema())


# ---------------------------------------------------------------------------
# find schemas (Task 8)
# ---------------------------------------------------------------------------


def test_find_payload_snapshot(snapshot: Any) -> None:
    snapshot("find_payload", FindPayload.model_json_schema())


def test_find_results_snapshot(snapshot: Any) -> None:
    snapshot("find_results", FindResults.model_json_schema())


def test_email_hit_snapshot(snapshot: Any) -> None:
    snapshot("email_hit", EmailHit.model_json_schema())


def test_file_hit_snapshot(snapshot: Any) -> None:
    snapshot("file_hit", FileHit.model_json_schema())


def test_message_hit_snapshot(snapshot: Any) -> None:
    snapshot("message_hit", MessageHit.model_json_schema())


def test_page_hit_snapshot(snapshot: Any) -> None:
    snapshot("page_hit", PageHit.model_json_schema())


def test_search_hit_union_snapshot(snapshot: Any) -> None:
    schema = TypeAdapter(SearchHit).json_schema()
    snapshot("search_hit_union", schema)

    # Gate: discriminated union on ``kind`` must render with both keywords
    on_disk = json.loads((_SNAPSHOT_DIR / "search_hit_union.json").read_text())
    assert "oneOf" in on_disk, "search_hit_union.json must contain 'oneOf'"
    assert "discriminator" in on_disk, "search_hit_union.json must contain 'discriminator'"


# ---------------------------------------------------------------------------
# whats_new schemas (Task 7)
# ---------------------------------------------------------------------------


def test_whats_new_payload(snapshot: Any) -> None:
    snapshot("whats_new_payload", WhatsNewPayload.model_json_schema())


def test_whats_new_summary(snapshot: Any) -> None:
    snapshot("whats_new_summary", WhatsNewSummary.model_json_schema())


def test_mail_item(snapshot: Any) -> None:
    snapshot("mail_item", MailItem.model_json_schema())


def test_whats_new_event_summary(snapshot: Any) -> None:
    snapshot("whats_new_event_summary", WhatsNewEventSummary.model_json_schema())


def test_teams_item(snapshot: Any) -> None:
    snapshot("teams_item", TeamsItem.model_json_schema())


# ---------------------------------------------------------------------------
# resolve schemas (Task 9)
# ---------------------------------------------------------------------------


def test_resolve_payload(snapshot: Any) -> None:
    snapshot("resolve_payload", ResolvePayload.model_json_schema())


def test_resolved_content(snapshot: Any) -> None:
    snapshot("resolved_content", ResolvedContent.model_json_schema())


def test_resolved_content_data_union(snapshot: Any) -> None:
    schema = TypeAdapter(ResolvedContentData).json_schema()
    snapshot("resolved_content_data_union", schema)

    # Gate: discriminated union on ``kind`` must render with both keywords
    on_disk = json.loads((_SNAPSHOT_DIR / "resolved_content_data_union.json").read_text())
    assert "oneOf" in on_disk, "resolved_content_data_union.json must contain 'oneOf'"
    assert "discriminator" in on_disk, "resolved_content_data_union.json must contain 'discriminator'"
    # Gate: must have exactly 7 variants
    assert len(on_disk["oneOf"]) == 7, (
        f"Expected 7 variants in resolved_content_data_union.json, got {len(on_disk['oneOf'])}"
    )
    assert on_disk["discriminator"]["propertyName"] == "kind", (
        "discriminator propertyName must be 'kind'"
    )


def test_email_content(snapshot: Any) -> None:
    snapshot("email_content", EmailContent.model_json_schema())


def test_chat_thread_content(snapshot: Any) -> None:
    snapshot("chat_thread_content", ChatThreadContent.model_json_schema())


def test_chat_message_content(snapshot: Any) -> None:
    snapshot("chat_message_content", ChatMessageContent.model_json_schema())


def test_channel_message_content(snapshot: Any) -> None:
    snapshot("channel_message_content", ChannelMessageContent.model_json_schema())


def test_meeting_content(snapshot: Any) -> None:
    snapshot("meeting_content", MeetingContent.model_json_schema())


def test_sharepoint_page_content(snapshot: Any) -> None:
    snapshot("sharepoint_page_content", SharePointPageContent.model_json_schema())


def test_onedrive_file_content(snapshot: Any) -> None:
    snapshot("onedrive_file_content", OneDriveFileContent.model_json_schema())
