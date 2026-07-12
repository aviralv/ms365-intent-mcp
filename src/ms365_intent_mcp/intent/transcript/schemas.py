"""Schemas for transcript — download a meeting recording's VTT transcript.

Payload accepts ``url`` OR ``name`` (at least one). ``url`` is the fast path
(a recording URL, e.g. ``meeting()``'s ``vroom_url``); ``name`` triggers
filename-based discovery across own-drive, Graph Search, and Teams chat events.
"""

from __future__ import annotations

from typing import Annotated, Literal

from pydantic import BaseModel, ConfigDict, Field, model_validator

from .._shared import BaseResponse


class TranscriptPayload(BaseModel):
    """Payload for transcript. Provide ``url`` or ``name`` (at least one)."""

    model_config = ConfigDict(extra="forbid")
    url: Annotated[
        str | None,
        Field(
            default=None,
            description=(
                "A recording URL — Teams recap link, SharePoint share/sharable "
                "link, or meeting()'s vroom_url. Fast path when it carries "
                "drive/item IDs."
            ),
        ),
    ]
    name: Annotated[
        str | None,
        Field(
            default=None,
            description=(
                "Meeting name (or substring) to discover the recording by "
                "filename across own-drive, Graph Search, and Teams chats."
            ),
        ),
    ]
    output_dir: Annotated[
        str | None,
        Field(
            default=None,
            description=(
                "Directory to write the VTT to. Defaults to "
                "~/.cache/ms365-intent-mcp/transcripts."
            ),
        ),
    ]

    @model_validator(mode="after")
    def _require_url_or_name(self) -> "TranscriptPayload":
        if not (self.url or self.name):
            raise ValueError("Provide either `url` or `name`.")
        return self


class TranscriptResultResponse(BaseResponse):
    """Response for transcript — where the VTT landed plus meeting metadata."""

    type: Literal["transcript_result"] = "transcript_result"
    status: Literal["ok", "error"]
    file_path: str | None = None
    meeting_name: str | None = None
    meeting_date: str | None = None
    line_count: int = 0
    byte_count: int = 0
    has_speaker_tags: bool = False
    message: str | None = None
    rendered_markdown: str
