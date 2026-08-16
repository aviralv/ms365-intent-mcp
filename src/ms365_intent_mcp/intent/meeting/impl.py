"""meeting implementation — wraps composers.meeting.compose_meeting."""

from __future__ import annotations

from fastmcp import Context

from ...composers.meeting import compose_meeting
from .._helpers import _get_deps, wrap_errors
from .schemas import (
    AttendeeStatus,
    MeetingDetail,
    MeetingPayload,
    OnlineMeetingInfo,
    PersonRef,
    RecordingMetadata,
)
from datetime import UTC

TOOL_NAME = "meeting"


@wrap_errors(TOOL_NAME)
async def _meeting_impl(ctx: Context, payload: MeetingPayload) -> MeetingDetail:
    """Call the underlying composer and return a typed MeetingDetail."""
    config, client, permissions = _get_deps(ctx)
    data, markdown = await compose_meeting(
        client,
        permissions,
        payload.identifier,
        config.default_timezone,
    )

    organizer_data = data.get("organizer") or {}
    organizer = PersonRef(
        name=organizer_data.get("name") or "unknown",
        email=organizer_data.get("email"),
    )

    attendees = []
    for a in data.get("attendees", []):
        try:
            attendees.append(AttendeeStatus.model_validate(a))
        except Exception:
            pass

    online_meeting = None
    if data.get("online_meeting"):
        try:
            online_meeting = OnlineMeetingInfo.model_validate(data["online_meeting"])
        except Exception:
            pass

    recording = None
    if data.get("recording"):
        try:
            recording = RecordingMetadata.model_validate(data["recording"])
        except Exception:
            pass

    # Parse datetimes — use a fallback for empty strings
    from datetime import datetime
    from datetime import timezone as _tz

    def _parse_dt(s: str):
        if not s:
            return datetime.now(UTC)
        try:
            return datetime.fromisoformat(s.replace("Z", "+00:00"))
        except (ValueError, AttributeError):
            return datetime.now(UTC)

    return MeetingDetail(
        id=data.get("id") or "unknown",
        subject=data.get("subject") or payload.identifier,
        start=_parse_dt(data.get("start", "")),
        end=_parse_dt(data.get("end", "")),
        organizer=organizer,
        attendees=attendees,
        location=data.get("location"),
        online_meeting=online_meeting,
        recording=recording,
        rendered_markdown=markdown,
    )
