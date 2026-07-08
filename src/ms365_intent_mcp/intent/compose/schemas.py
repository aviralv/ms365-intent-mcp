"""Schemas for compose — discriminated union on ``type``.

Three variants:
- ``ComposeEmail`` (mode='new' | 'reply' | 'reply_all' | 'forward')
- ``ComposeEvent``
- ``ComposeTeamsMessage``

Each has its own response model. ``ComposePayload`` is the union used
by the FastMCP tool registration.
"""

from __future__ import annotations

from datetime import datetime
from typing import Annotated, Literal, Union

from pydantic import BaseModel, ConfigDict, Field, HttpUrl, model_validator

from .._shared import Attendee, BaseResponse, Recipient


# ============================================================
# Payload variants
# ============================================================


class ComposeEmail(BaseModel):
    """Draft an email — new, reply, reply_all, or forward.

    Mode-specific field requirements are enforced by ``_check_mode_fields``:
    - ``mode='new'`` requires ``to`` and ``subject``
    - ``mode='reply' | 'reply_all' | 'forward'`` requires ``in_reply_to_message_id``
    """

    model_config = ConfigDict(extra="forbid")
    type: Literal["email"]
    mode: Literal["new", "reply", "reply_all", "forward"] = "new"
    body: Annotated[str, Field(min_length=1, description="HTML supported.")]
    # new / forward:
    to: list[Recipient] | None = None
    cc: list[Recipient] | None = None
    subject: str | None = None
    importance: Literal["low", "normal", "high"] = "normal"
    # reply / reply_all / forward:
    in_reply_to_message_id: str | None = None
    comment: str | None = None
    idempotency_key: str | None = Field(default=None, min_length=1)

    @model_validator(mode="after")
    def _check_mode_fields(self) -> "ComposeEmail":
        if self.mode == "forward":
            raise ValueError("mode='forward' is not yet supported — will land in v0.8.1")
        needs_parent = self.mode in ("reply", "reply_all", "forward")
        needs_new = self.mode == "new"
        if needs_parent and not self.in_reply_to_message_id:
            raise ValueError(f"mode='{self.mode}' requires in_reply_to_message_id")
        if needs_new and not (self.to and self.subject):
            raise ValueError("mode='new' requires both 'to' and 'subject'")
        return self


class ComposeEvent(BaseModel):
    """Create a calendar event.

    Validates ``end > start`` and duration ≤ 12 hours. Longer events almost
    always indicate an off-by-day error in the caller's timezone handling.
    """

    model_config = ConfigDict(extra="forbid")
    type: Literal["event"]
    subject: Annotated[str, Field(min_length=1)]
    start: datetime
    end: datetime
    timezone: str = Field(
        description="IANA timezone name (e.g. 'Europe/Berlin'). Interprets naive start/end.",
    )
    attendees: list[Attendee] | None = None
    location: str | None = None
    body: str | None = None
    is_online_meeting: bool = False
    idempotency_key: str | None = Field(default=None, min_length=1)

    @model_validator(mode="after")
    def _check_times(self) -> "ComposeEvent":
        if self.end <= self.start:
            raise ValueError("end must be after start")
        duration_minutes = (self.end - self.start).total_seconds() / 60
        if duration_minutes > 60 * 12:
            raise ValueError("event duration exceeds 12 hours")
        return self


class ComposeTeamsMessage(BaseModel):
    """Send a Teams chat message."""

    model_config = ConfigDict(extra="forbid")
    type: Literal["teams_message"]
    chat_id: Annotated[str, Field(min_length=1)]
    content: Annotated[str, Field(min_length=1)]
    content_type: Literal["text", "html"] = "text"
    idempotency_key: str | None = Field(default=None, min_length=1)


ComposePayload = Annotated[
    Union[ComposeEmail, ComposeEvent, ComposeTeamsMessage],
    Field(discriminator="type"),
]


# ============================================================
# Response variants
# ============================================================


class EmailDraftCreated(BaseResponse):
    """Response: an email draft was saved to the user's Drafts folder."""

    type: Literal["email_draft_created"] = "email_draft_created"
    draft_id: str
    subject: str
    to: list[Recipient]
    web_link: HttpUrl
    rendered_markdown: str


class EventCreated(BaseResponse):
    """Response: a calendar event was created."""

    type: Literal["event_created"] = "event_created"
    event_id: str
    subject: str
    start: datetime
    end: datetime
    join_url: HttpUrl | None = None
    rendered_markdown: str


class TeamsMessageSent(BaseResponse):
    """Response: a Teams chat message was sent."""

    type: Literal["teams_message_sent"] = "teams_message_sent"
    message_id: str
    chat_id: str
    rendered_markdown: str


ComposeResponse = Union[EmailDraftCreated, EventCreated, TeamsMessageSent]
