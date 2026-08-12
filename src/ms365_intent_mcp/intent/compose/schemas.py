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
from typing import Annotated, Literal

from pydantic import BaseModel, ConfigDict, Field, HttpUrl, model_validator

from .._shared import Attendee, BaseResponse, Recipient


# ============================================================
# Payload variants
# ============================================================


class ComposeEmail(BaseModel):
    """Draft an email — new, reply, reply_all, or forward."""

    model_config = ConfigDict(extra="forbid")
    type: Literal["email"]
    mode: Literal["new", "reply", "reply_all", "forward"] = "new"
    body: Annotated[str, Field(min_length=1, description="HTML supported.")]
    # new / forward:
    to: list[Recipient] | None = None
    cc: list[Recipient] | None = None
    # new only:
    subject: str | None = None
    importance: Literal["low", "normal", "high"] = "normal"
    # reply / reply_all / forward:
    in_reply_to_message_id: str | None = None
    # reply / reply_all only:
    comment: str | None = None
    idempotency_key: str | None = Field(default=None, min_length=1)

    @model_validator(mode="after")
    def _check_mode_fields(self) -> "ComposeEmail":
        needs_parent = self.mode in ("reply", "reply_all", "forward")
        if needs_parent and not self.in_reply_to_message_id:
            raise ValueError(f"mode='{self.mode}' requires in_reply_to_message_id")
        if self.mode == "forward" and not self.to:
            raise ValueError("mode='forward' requires 'to' (at least one recipient)")
        if self.mode == "new" and not (self.to and self.subject):
            raise ValueError("mode='new' requires both 'to' and 'subject'")
        return self


class ComposeEvent(BaseModel):
    """Create a calendar event (mode='create'), or forward an existing one to
    new recipients (mode='forward', which sends immediately — no draft)."""

    model_config = ConfigDict(extra="forbid")
    type: Literal["event"]
    mode: Literal["create", "forward"] = "create"
    subject: str | None = None
    start: datetime | None = None
    end: datetime | None = None
    timezone: str | None = Field(
        default=None,
        description="IANA timezone (e.g. 'Europe/Berlin'); required for mode='create'.",
    )
    attendees: list[Attendee] | None = None
    location: str | None = None
    body: str | None = None
    is_online_meeting: bool = False
    # forward mode:
    event_id: str | None = None
    to: list[Recipient] | None = None
    comment: str | None = None
    idempotency_key: str | None = Field(default=None, min_length=1)

    @model_validator(mode="after")
    def _check_times(self) -> "ComposeEvent":
        if self.mode == "forward":
            if not self.event_id:
                raise ValueError("mode='forward' requires 'event_id'")
            if not self.to:
                raise ValueError("mode='forward' requires 'to' (at least one recipient)")
            if self.subject or self.start or self.end:
                raise ValueError(
                    "mode='forward' does not accept subject/start/end — "
                    "forwarding sends the existing invite unchanged"
                )
            return self
        # create mode
        if not (self.subject and self.start and self.end and self.timezone):
            raise ValueError("mode='create' requires subject, start, end, and timezone")
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
    ComposeEmail | ComposeEvent | ComposeTeamsMessage,
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


class EventForwarded(BaseResponse):
    """Response: a calendar event was forwarded (sent immediately, 202 — no body)."""

    type: Literal["event_forwarded"] = "event_forwarded"
    to: list[Recipient]
    rendered_markdown: str


class TeamsMessageSent(BaseResponse):
    """Response: a Teams chat message was sent."""

    type: Literal["teams_message_sent"] = "teams_message_sent"
    message_id: str
    chat_id: str
    rendered_markdown: str


ComposeResponse = EmailDraftCreated | EventCreated | EventForwarded | TeamsMessageSent
