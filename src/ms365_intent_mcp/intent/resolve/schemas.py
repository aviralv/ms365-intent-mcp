"""Schemas for resolve — flat payload, discriminated union on ``kind`` for response.

Input:
- ``ResolvePayload`` — a single ``url`` field. No discriminator on input: the
  URL type is auto-detected by the regex dispatch table in resolver.py. Adding
  a url_type discriminator was considered and rejected (gives the LLM one more
  field to get wrong with no upside).

Output:
- ``ResolvedContent`` — a ``BaseResponse`` with ``kind`` pinned to the detected
  URL type and ``data`` as a 7-variant discriminated union on ``kind``.

Structured fields on content types are placeholder stubs until Task 12
refactors composers to return ``(dict, markdown)`` tuples.
"""

from __future__ import annotations

from datetime import datetime
from typing import Annotated, Literal, Union

from pydantic import BaseModel, ConfigDict, Field, HttpUrl

from .._shared import BaseResponse


# ============================================================
# Payload
# ============================================================


class ResolvePayload(BaseModel):
    """Resolve any M365 URL."""

    model_config = ConfigDict(extra="forbid")
    url: HttpUrl


# ============================================================
# Content type variants (discriminated on ``kind``)
# ============================================================


class EmailContent(BaseModel):
    """Content for a resolved email message."""

    model_config = ConfigDict(extra="forbid")
    kind: Literal["email"]
    subject: str
    sender: str
    body: str


class ChatThreadContent(BaseModel):
    """Content for a resolved Teams chat thread."""

    model_config = ConfigDict(extra="forbid")
    kind: Literal["chat_thread"]
    topic: str | None = None
    member_count: int = 0
    recent_message_count: int = 0


class ChatMessageContent(BaseModel):
    """Content for a resolved individual Teams chat message."""

    model_config = ConfigDict(extra="forbid")
    kind: Literal["chat_message"]
    sender: str
    body: str
    created: datetime | None = None


class ChannelMessageContent(BaseModel):
    """Content for a resolved Teams channel message."""

    model_config = ConfigDict(extra="forbid")
    kind: Literal["channel_message"]
    sender: str
    body: str
    channel_name: str | None = None


class MeetingContent(BaseModel):
    """Content for a resolved Teams meeting link."""

    model_config = ConfigDict(extra="forbid")
    kind: Literal["meeting"]
    subject: str
    start: datetime | None = None
    end: datetime | None = None


class SharePointPageContent(BaseModel):
    """Content for a resolved SharePoint page."""

    model_config = ConfigDict(extra="forbid")
    kind: Literal["sharepoint_page"]
    title: str
    web_url: HttpUrl | None = None


class OneDriveFileContent(BaseModel):
    """Content for a resolved OneDrive file."""

    model_config = ConfigDict(extra="forbid")
    kind: Literal["onedrive_file"]
    name: str
    web_url: HttpUrl | None = None
    size: int | None = None


ResolvedContentData = Annotated[
    Union[
        EmailContent,
        ChatThreadContent,
        ChatMessageContent,
        ChannelMessageContent,
        MeetingContent,
        SharePointPageContent,
        OneDriveFileContent,
    ],
    Field(discriminator="kind"),
]


# ============================================================
# Response
# ============================================================


class ResolvedContent(BaseResponse):
    """Response for resolve."""

    type: Literal["resolved_content"] = "resolved_content"
    url: HttpUrl
    kind: Literal[
        "email",
        "chat_thread",
        "chat_message",
        "channel_message",
        "meeting",
        "sharepoint_page",
        "onedrive_file",
    ]
    data: ResolvedContentData
    rendered_markdown: str
