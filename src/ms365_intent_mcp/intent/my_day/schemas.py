"""Schemas for my_day — single payload, structured summary response.

Structured fields (events, mail, teams) are stubs until Task 12 refactors
composers to return ``(dict, markdown)`` tuples. ``rendered_markdown``
carries the current composer output verbatim.
"""

from __future__ import annotations

import datetime as _dt
from typing import Literal

from pydantic import BaseModel, ConfigDict, Field

from .._shared import BaseResponse


class MyDayPayload(BaseModel):
    """Payload for my_day."""

    model_config = ConfigDict(extra="forbid")
    date: _dt.date | None = Field(default=None, description="Defaults to today (server tz).")


class EventSummary(BaseModel):
    """Compact summary of a single calendar event."""

    model_config = ConfigDict(extra="forbid")
    subject: str
    start: _dt.datetime
    end: _dt.datetime
    location: str | None = None
    is_online_meeting: bool = False


class MailSummary(BaseModel):
    """Aggregate mail counts for the day."""

    model_config = ConfigDict(extra="forbid")
    unread_count: int = 0
    relevant_count: int = 0
    flagged_count: int = 0


class TeamsActivitySummary(BaseModel):
    """Aggregate Teams activity for the day."""

    model_config = ConfigDict(extra="forbid")
    recent_message_count: int = 0


class MyDaySummary(BaseResponse):
    """Response for my_day."""

    type: Literal["my_day_summary"] = "my_day_summary"
    date: _dt.date
    events: list[EventSummary]
    mail: MailSummary
    teams: TeamsActivitySummary
    rendered_markdown: str
