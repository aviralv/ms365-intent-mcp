"""Schemas for whats_new — payload + structured summary response.

Structured fields (mail, events, teams) are stubs until Task 12 refactors
composers to return ``(dict, markdown)`` tuples. ``rendered_markdown``
carries the current composer output verbatim.
"""

from __future__ import annotations

from datetime import datetime
from typing import Literal

from pydantic import BaseModel, ConfigDict

from .._shared import BaseResponse


class WhatsNewPayload(BaseModel):
    """Payload for whats_new."""

    model_config = ConfigDict(extra="forbid")
    since: datetime  # ISO 8601; naive treated as UTC per Pydantic default
    scope: Literal["mail", "calendar", "teams", "all"] = "all"


class EventSummary(BaseModel):
    """Compact summary of a single calendar event."""

    model_config = ConfigDict(extra="forbid")
    subject: str
    start: datetime
    end: datetime
    location: str | None = None
    is_online_meeting: bool = False


class MailItem(BaseModel):
    """Single mail item summary."""

    model_config = ConfigDict(extra="forbid")
    subject: str
    sender: str
    received: datetime
    is_read: bool = False
    importance: Literal["low", "normal", "high"] = "normal"


class TeamsItem(BaseModel):
    """Single Teams message summary."""

    model_config = ConfigDict(extra="forbid")
    chat_name: str | None = None
    sender: str
    body_preview: str
    received: datetime


class WhatsNewSummary(BaseResponse):
    """Response for whats_new."""

    type: Literal["whats_new_summary"] = "whats_new_summary"
    since: datetime
    mail: list[MailItem]
    events: list[EventSummary]
    teams: list[TeamsItem]
    rendered_markdown: str
