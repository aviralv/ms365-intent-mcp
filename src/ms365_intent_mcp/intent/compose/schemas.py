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
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

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
    def _check_mode_fields(self) -> ComposeEmail:
        needs_parent = self.mode in ("reply", "reply_all", "forward")
        if needs_parent and not self.in_reply_to_message_id:
            raise ValueError(f"mode='{self.mode}' requires in_reply_to_message_id")
        if self.mode == "forward" and not self.to:
            raise ValueError("mode='forward' requires 'to' (at least one recipient)")
        if self.mode == "new" and not (self.to and self.subject):
            raise ValueError("mode='new' requires both 'to' and 'subject'")
        return self


class ComposeEvent(BaseModel):
    """Create, forward (sent immediately), or partially update an existing event.

    Update supports only subject/start/end/location; omitted fields are preserved.
    """

    model_config = ConfigDict(extra="forbid")
    type: Literal["event"]
    mode: Literal["create", "forward", "update"] = "create"
    subject: str | None = None
    start: datetime | None = None
    end: datetime | None = None
    timezone: str | None = Field(
        default=None,
        description="IANA timezone; required for create and naive update timestamps. "
        "Aware update timestamps are converted to this zone (UTC when omitted).",
    )
    attendees: list[Attendee] | None = None
    location: str | None = None
    body: str | None = None
    is_online_meeting: bool = False
    # forward/update modes:
    event_id: str | None = Field(default=None, description="Exact Graph event/occurrence ID.")
    to: list[Recipient] | None = None
    comment: str | None = None
    idempotency_key: str | None = Field(default=None, min_length=1)

    @model_validator(mode="after")
    def _check_times(self) -> ComposeEvent:
        if self.mode == "update":
            if not self.event_id or not self.event_id.strip():
                raise ValueError("mode='update' requires 'event_id'")
            changes = self.model_fields_set & {"subject", "start", "end", "location"}
            if not changes:
                raise ValueError("mode='update' requires subject, start, end, or location")
            unsupported = self.model_fields_set & {
                "body",
                "attendees",
                "is_online_meeting",
                "to",
                "comment",
            }
            if unsupported:
                raise ValueError(f"mode='update' does not accept {', '.join(sorted(unsupported))}")
            if any(getattr(self, field) is None for field in changes):
                raise ValueError("update fields cannot be null; omit fields to preserve them")
            if "subject" in changes and not self.subject.strip():
                raise ValueError("updated subject cannot be empty")
            if "timezone" in self.model_fields_set:
                try:
                    ZoneInfo(self.timezone or "")
                except (ZoneInfoNotFoundError, ValueError) as exc:
                    raise ValueError("timezone must be a valid IANA timezone") from exc
                if not changes & {"start", "end"}:
                    raise ValueError("timezone requires an updated start or end")
            if (
                any(value is not None and value.tzinfo is None for value in (self.start, self.end))
                and not self.timezone
            ):
                raise ValueError("naive update timestamps require an explicit IANA timezone")
            # Compare instants after timezone/DST normalization in the composer,
            # including the stored endpoint when only one endpoint is supplied.
            return self
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


class FileAttachment(BaseModel):
    """A reference attachment (link to an existing OneDrive/SharePoint file)."""

    model_config = ConfigDict(extra="forbid")
    name: Annotated[str, Field(min_length=1, description="Display name (e.g. 'report.xlsx').")]
    url: Annotated[str, Field(min_length=1, description="SharePoint or OneDrive URL of the file.")]


class ComposeTeamsMessage(BaseModel):
    """Send a Teams chat message, optionally with file attachments."""

    model_config = ConfigDict(extra="forbid")
    type: Literal["teams_message"]
    chat_id: Annotated[str, Field(min_length=1)]
    content: Annotated[str, Field(min_length=1)]
    content_type: Literal["text", "html"] = "text"
    attachments: list[FileAttachment] | None = Field(
        default=None,
        description="Optional file attachments (OneDrive/SharePoint URLs) to include.",
    )
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


class EventUpdated(BaseResponse):
    """Response: a targeted event was updated; times include explicit offsets."""

    type: Literal["event_updated"] = "event_updated"
    event_id: str
    subject: str
    start: datetime
    end: datetime
    timezone: str
    web_link: HttpUrl | None = None
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


ComposeResponse = (
    EmailDraftCreated | EventCreated | EventForwarded | EventUpdated | TeamsMessageSent
)
