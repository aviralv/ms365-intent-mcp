"""FastMCP registration for schedule."""

from __future__ import annotations

from typing import Annotated

from fastmcp import Context, FastMCP

from .impl import _schedule_impl
from .schemas import SchedulePayload, ScheduleSuggestions


def register(mcp: FastMCP) -> None:
    """Register the schedule tool on a FastMCP instance."""

    @mcp.tool()
    async def schedule(
        ctx: Context,
        payload: Annotated[
            SchedulePayload,
            "Find meeting times. attendees (min 1) must include email; "
            "duration_minutes defaults to 30 (5–480). "
            "Optional constraints.start/end narrow the search window; "
            "constraints.working_hours_only (default true) limits to business hours.",
        ],
    ) -> ScheduleSuggestions:
        """Find available meeting times. Returns ranked time slots with confidence scores."""
        return await _schedule_impl(ctx, payload)
