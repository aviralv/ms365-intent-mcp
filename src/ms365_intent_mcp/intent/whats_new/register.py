"""FastMCP registration for whats_new_v1."""

from __future__ import annotations

from typing import Annotated

from fastmcp import Context, FastMCP

from .impl import _whats_new_v1_impl
from .schemas import WhatsNewPayload, WhatsNewSummary


def register(mcp: FastMCP) -> None:
    """Register the whats_new_v1 tool on a FastMCP instance."""

    @mcp.tool()
    async def whats_new_v1(
        ctx: Context,
        payload: Annotated[
            WhatsNewPayload,
            "What happened since a given time. "
            "since: ISO 8601 datetime (required). "
            "scope: 'mail', 'calendar', 'teams', or 'all' (default).",
        ],
    ) -> WhatsNewSummary:
        """What happened since a given time? Returns new mail, events, and Teams messages."""
        return await _whats_new_v1_impl(ctx, payload)
