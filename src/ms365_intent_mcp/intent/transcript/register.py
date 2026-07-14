"""FastMCP registration for transcript."""

from __future__ import annotations

from typing import Annotated

from fastmcp import Context, FastMCP

from .impl import _transcript_impl
from .schemas import TranscriptPayload, TranscriptResultResponse


def register(mcp: FastMCP) -> None:
    """Register the transcript tool on a FastMCP instance."""

    @mcp.tool()
    async def transcript(
        ctx: Context,
        payload: Annotated[
            TranscriptPayload,
            "Download a meeting recording's VTT transcript, or list recordings. "
            "Provide `url` (a recording URL / meeting()'s vroom_url — fast "
            "path), `name` (discover by meeting name), an "
            "`item_id`+`drive_id`+`site_root` triple (deterministic, zero "
            "discovery), or `list=true` (enumerate recordings newest-first — "
            "use to find ad-hoc calls that name-search can't match). Optional "
            "`output_dir`.",
        ],
    ) -> TranscriptResultResponse:
        """Get the VTT transcript for a meeting recording (or list recordings). Writes the file to disk and returns the path."""
        return await _transcript_impl(ctx, payload)
