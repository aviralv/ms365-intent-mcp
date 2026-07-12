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
            "Download a meeting recording's VTT transcript. Provide `url` "
            "(a recording URL / meeting()'s vroom_url — fast path) or `name` "
            "(discover by meeting name). Optional `output_dir`.",
        ],
    ) -> TranscriptResultResponse:
        """Get the VTT transcript for a meeting recording. Writes the file to disk and returns the path."""
        return await _transcript_impl(ctx, payload)
