"""FastMCP registration for meeting."""

from __future__ import annotations

from typing import Annotated

from fastmcp import Context, FastMCP

from .impl import _meeting_impl
from .schemas import MeetingDetail, MeetingPayload


def register(mcp: FastMCP) -> None:
    """Register the meeting tool on a FastMCP instance."""

    @mcp.tool()
    async def meeting(
        ctx: Context,
        payload: Annotated[
            MeetingPayload,
            "Full context for a meeting. identifier can be an event ID, "
            "a subject substring, or 'next' for the upcoming meeting.",
        ],
    ) -> MeetingDetail:
        """Tell me about this meeting. Returns full context: attendees, body, Teams link, recording."""
        return await _meeting_impl(ctx, payload)
