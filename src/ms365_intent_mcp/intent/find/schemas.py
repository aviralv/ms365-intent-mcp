from __future__ import annotations

from datetime import datetime
from typing import Annotated, Literal

from pydantic import BaseModel, ConfigDict, Field, HttpUrl

from .._shared import BaseResponse


class FindPayload(BaseModel):
    model_config = ConfigDict(extra="forbid")
    query: Annotated[str, Field(min_length=1)]
    entity_type: Literal["email", "file", "message", "page"] | None = Field(
        default=None,
        description="Optional filter — restricts search to one entity type. Not a discriminator.",
    )


class EmailHit(BaseModel):
    model_config = ConfigDict(extra="forbid")
    kind: Literal["email"]
    subject: str
    sender: str
    body_preview: str
    web_link: HttpUrl | None = None


class FileHit(BaseModel):
    model_config = ConfigDict(extra="forbid")
    kind: Literal["file"]
    name: str
    web_url: HttpUrl | None = None
    size: int | None = None


class MessageHit(BaseModel):
    model_config = ConfigDict(extra="forbid")
    kind: Literal["message"]
    sender: str
    body_preview: str
    created: datetime | None = None
    chat_id: str = ""
    chat_url: str = ""


class PageHit(BaseModel):
    model_config = ConfigDict(extra="forbid")
    kind: Literal["page"]
    title: str
    web_url: HttpUrl | None = None


SearchHit = Annotated[
    EmailHit | FileHit | MessageHit | PageHit,
    Field(discriminator="kind"),
]


class FindResults(BaseResponse):
    type: Literal["find_results"] = "find_results"
    query: str
    hits: list[SearchHit]
    rendered_markdown: str
