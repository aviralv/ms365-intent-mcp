"""Schemas for meeting — single payload, structured detail response.

Structured fields (id, subject, start, end, organizer, attendees) are stubs
until Task 12 refactors composers to return ``(dict, markdown)`` tuples.
``rendered_markdown`` carries the current composer output verbatim.
"""

from __future__ import annotations

from datetime import date, datetime
from typing import Annotated, Literal

from pydantic import BaseModel, ConfigDict, EmailStr, Field, HttpUrl

from .._shared import BaseResponse


class MeetingPayload(BaseModel):
    """Payload for meeting."""

    model_config = ConfigDict(extra="forbid")
    identifier: Annotated[
        str,
        Field(min_length=1, description="Event ID, subject substring, or 'next'."),
    ]


class PersonRef(BaseModel):
    """Lightweight reference to a person (organizer, etc.)."""

    model_config = ConfigDict(extra="forbid")
    name: str
    email: EmailStr | None = None


class AttendeeStatus(BaseModel):
    """One attendee with their RSVP response."""

    model_config = ConfigDict(extra="forbid")
    name: str
    email: EmailStr | None = None
    response: Literal["accepted", "declined", "tentative", "none", "notResponded"] = "none"


class OnlineMeetingInfo(BaseModel):
    """Teams join link for an online meeting."""

    model_config = ConfigDict(extra="forbid")
    join_url: HttpUrl


class RecordingMetadata(BaseModel):
    """Recording metadata surfaced when a meeting has a recording (issue #8)."""

    model_config = ConfigDict(extra="forbid")
    recording_url: HttpUrl
    display_name: str | None = None
    transcript_ready: bool = False
    drive_id: str | None = None
    drive_item_id: str | None = None
    owner_upn: str | None = None
    vroom_url: HttpUrl | None = None


class MeetingDetail(BaseResponse):
    """Response for meeting."""

    type: Literal["meeting_detail"] = "meeting_detail"
    id: str
    subject: str
    start: date | datetime
    end: date | datetime
    start_timezone: str | None = None
    end_timezone: str | None = None
    organizer: PersonRef
    attendees: list[AttendeeStatus]
    location: str | None = None
    online_meeting: OnlineMeetingInfo | None = None
    recording: RecordingMetadata | None = None
    rendered_markdown: str
