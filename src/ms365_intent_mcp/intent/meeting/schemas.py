"""Schemas for meeting — single payload, structured detail response.

Structured fields (id, subject, start, end, organizer, attendees) are stubs
until Task 12 refactors composers to return ``(dict, markdown)`` tuples.
``rendered_markdown`` carries the current composer output verbatim.
"""

from __future__ import annotations

import re
from datetime import date as calendar_date
from datetime import datetime
from typing import Annotated, Literal

from pydantic import (
    BaseModel,
    ConfigDict,
    EmailStr,
    Field,
    HttpUrl,
    field_validator,
    model_validator,
)

from .._shared import BaseResponse


class MeetingPayload(BaseModel):
    """Payload for meeting."""

    model_config = ConfigDict(extra="forbid")
    identifier: Annotated[
        str,
        Field(min_length=1, description="Event ID, subject substring, or 'next'."),
    ]
    date: calendar_date | None = Field(
        default=None,
        description="Occurrence start date (YYYY-MM-DD) in the configured calendar timezone. "
        "With an event ID, validates that occurrence; series-master IDs are rejected. "
        "Omit to retain legacy lookup. Not supported with 'next'.",
    )

    @field_validator("date", mode="before")
    @classmethod
    def validate_date(cls, value):
        if value is None or type(value) is calendar_date:
            return value
        if not isinstance(value, str) or not re.fullmatch(r"\d{4}-\d{2}-\d{2}", value):
            raise ValueError("date must be YYYY-MM-DD")
        return calendar_date.fromisoformat(value)

    @model_validator(mode="after")
    def validate_next(self):
        if self.date is not None and self.identifier.lower() == "next":
            raise ValueError("date cannot be combined with 'next'; use a subject or occurrence ID")
        return self


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
    # Occurrence date-matching (issue #79/#51): recording_date is the recording's
    # own date; date_matches_occurrence is False when the chosen recording is
    # from a different occurrence than requested (None when the occurrence date
    # couldn't be determined).
    recording_date: str | None = None
    occurrence_date: str | None = None
    date_matches_occurrence: bool | None = None


class MeetingDetail(BaseResponse):
    """Response for meeting."""

    type: Literal["meeting_detail"] = "meeting_detail"
    id: str
    subject: str
    start: calendar_date | datetime
    end: calendar_date | datetime
    start_timezone: str | None = None
    end_timezone: str | None = None
    organizer: PersonRef
    attendees: list[AttendeeStatus]
    location: str | None = None
    online_meeting: OnlineMeetingInfo | None = None
    recording: RecordingMetadata | None = None
    rendered_markdown: str
