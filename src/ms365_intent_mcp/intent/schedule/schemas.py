"""schedule_v1 schemas — payload, constraints, and response types."""

from __future__ import annotations

from datetime import datetime
from typing import Annotated, Literal

from pydantic import BaseModel, ConfigDict, Field

from .._shared import Attendee, BaseResponse


class ScheduleConstraints(BaseModel):
    model_config = ConfigDict(extra="forbid")
    start: datetime | None = None
    end: datetime | None = None
    working_hours_only: bool = True


class SchedulePayload(BaseModel):
    model_config = ConfigDict(extra="forbid")
    attendees: Annotated[list[Attendee], Field(min_length=1)]
    duration_minutes: Annotated[int, Field(default=30, ge=5, le=480)]
    constraints: ScheduleConstraints | None = None


class TimeSlot(BaseModel):
    model_config = ConfigDict(extra="forbid")
    start: datetime
    end: datetime
    confidence: float  # 0.0-1.0 from Graph findMeetingTimes


class ScheduleSuggestions(BaseResponse):
    type: Literal["schedule_suggestions"] = "schedule_suggestions"
    suggestions: list[TimeSlot]
    rendered_markdown: str
