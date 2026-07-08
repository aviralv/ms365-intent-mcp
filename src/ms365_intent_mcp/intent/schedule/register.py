"""FastMCP registration for schedule_v1."""

from __future__ import annotations

from typing import Annotated

from fastmcp import Context, FastMCP

from .impl import _schedule_v1_impl
from .schemas import SchedulePayload, ScheduleSuggestions


def register(mcp: FastMCP) -> None:
    """Register the schedule_v1 tool on a FastMCP instance."""

    @mcp.tool()
    async def schedule_v1(
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
        return await _schedule_v1_impl(ctx, payload)
