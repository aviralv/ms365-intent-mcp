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
            "Download a meeting recording's VTT, or list recordings. Provide "
            "exactly one: `url` (recording / meeting()'s vroom_url — fast path), "
            "`name` (discover by meeting name), `item_id`+`drive_id`+`site_root` "
            "(deterministic), or `list=true` (enumerate newest-first). Optional "
            "`output_dir`.",
        ],
    ) -> TranscriptResultResponse:
        """Get the VTT transcript for a meeting recording (or list recordings).

        Writes the file to disk and returns the path.
        """
        return await _transcript_impl(ctx, payload)
