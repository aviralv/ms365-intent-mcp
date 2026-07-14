"""Schemas for transcript — download a meeting recording's VTT transcript.

Input modes (mutually exclusive; provide exactly one):
  * ``url`` — a recording URL (fast path; e.g. ``meeting()``'s ``vroom_url``).
  * ``name`` — filename discovery across own-drive, Graph Search, and Teams
    chat events, then best-match. On multiple matches the freshest wins and the
    others are surfaced (issue #34).
  * ``item_id`` + ``drive_id`` + ``site_root`` — deterministic by-coords
    download, zero discovery (issue #33).
  * ``list=True`` — enumerate discovered recordings (newest first) instead of
    downloading. Ferret ``list`` parity; the escape hatch for ad-hoc calls
    titled from the other participant's side, which name-search can't match.
"""

from __future__ import annotations

from typing import Annotated, Literal

from pydantic import BaseModel, ConfigDict, Field, model_validator

from .._shared import BaseResponse


class TranscriptPayload(BaseModel):
    """Payload for transcript. Provide exactly one of: ``url``, ``name``, an
    ``item_id``+``drive_id``+``site_root`` triple, or ``list=True``."""

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
                "filename across own-drive, Graph Search, and Teams chats. On "
                "multiple matches, the most recent is downloaded and the others "
                "are surfaced. Note: ad-hoc 1:1 calls are titled from the "
                "recorder's side, so searching the counterpart's name may miss "
                "them — use list=true to find them by date."
            ),
        ),
    ]
    item_id: Annotated[
        str | None,
        Field(
            default=None,
            description=(
                "Drive-item id of a known recording — deterministic download "
                "with zero discovery. Requires drive_id and site_root too."
            ),
        ),
    ]
    drive_id: Annotated[
        str | None,
        Field(
            default=None,
            description="SharePoint drive id. Required with item_id.",
        ),
    ]
    site_root: Annotated[
        str | None,
        Field(
            default=None,
            description=(
                "SharePoint site root, e.g. "
                "https://{tenant}-my.sharepoint.com/personal/{user}. "
                "Required with item_id."
            ),
        ),
    ]
    list: Annotated[
        bool,
        Field(
            default=False,
            description=(
                "List discovered recordings (newest first) instead of "
                "downloading. Use to find a recording's id/name — especially "
                "ad-hoc calls that name-search can't match."
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
    def _require_exactly_one_input(self) -> "TranscriptPayload":
        coords = bool(self.item_id or self.drive_id or self.site_root)
        modes = [bool(self.url), bool(self.name), coords, self.list]
        if sum(modes) == 0:
            raise ValueError(
                "Provide one of: `url`, `name`, an item_id+drive_id+site_root "
                "triple, or `list=true`."
            )
        if sum(modes) > 1:
            raise ValueError(
                "`url`, `name`, the item_id triple, and `list` are mutually "
                "exclusive — provide exactly one."
            )
        if coords and not (self.item_id and self.drive_id and self.site_root):
            raise ValueError(
                "By-id download needs all three: `item_id`, `drive_id`, and "
                "`site_root`."
            )
        return self


class TranscriptResultResponse(BaseResponse):
    """Response for transcript — where the VTT landed plus meeting metadata,
    or (in list mode) the discovered recordings."""

    type: Literal["transcript_result"] = "transcript_result"
    status: Literal["ok", "error"]
    file_path: str | None = None
    meeting_name: str | None = None
    meeting_date: str | None = None
    line_count: int = 0
    byte_count: int = 0
    has_speaker_tags: bool = False
    alternatives_count: int = 0
    recordings: list[dict] | None = None
    message: str | None = None
    rendered_markdown: str
