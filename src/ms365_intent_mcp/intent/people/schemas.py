"""Schemas for people.

Single payload: ``PeoplePayload``.

Response model: ``PersonDetail`` (extends ``BaseResponse``).

Nested types: ``MailPreview``, ``ChatPreview``.

Note: ``EmailStr`` is available (pydantic[email] is installed as a
transitive dep via fastmcp). If it ever becomes unavailable, replace
``EmailStr`` with plain ``str`` throughout this file.
"""

from __future__ import annotations

from datetime import datetime
from typing import Annotated, Literal

from pydantic import BaseModel, ConfigDict, EmailStr, Field

from .._shared import BaseResponse


class PeoplePayload(BaseModel):
    model_config = ConfigDict(extra="forbid")
    query: Annotated[str, Field(min_length=1, description="Name or email substring.")]


class MailPreview(BaseModel):
    model_config = ConfigDict(extra="forbid")
    subject: str
    sender: str
    received: datetime | None = None


class ChatPreview(BaseModel):
    model_config = ConfigDict(extra="forbid")
    body: str
    last_message_at: datetime | None = None
    chat_id: str = ""
    chat_url: str = ""


class PersonDetail(BaseResponse):
    """Structured person lookup result.

    ``rendered_markdown`` carries the full legacy markdown from the composer.
    Structured fields (``name``, ``email``, ``job_title``, ``recent_mail``,
    ``recent_chat``) are populated with best-guess / placeholder values until
    Task 12 refactors the underlying composer to return structured data.
    """

    type: Literal["person_detail"] = "person_detail"
    name: str
    email: EmailStr | None = None
    job_title: str | None = None
    recent_mail: list[MailPreview]
    recent_chat: ChatPreview | None = None
    rendered_markdown: str
