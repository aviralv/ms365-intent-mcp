"""FastMCP registration for my_day_v1."""

from __future__ import annotations

from typing import Annotated

from fastmcp import Context, FastMCP

from .impl import _my_day_v1_impl
from .schemas import MyDayPayload, MyDaySummary


def register(mcp: FastMCP) -> None:
    """Register the my_day_v1 tool on a FastMCP instance."""

    @mcp.tool()
    async def my_day_v1(
        ctx: Context,
        payload: Annotated[MyDayPayload, "Daily overview: calendar, mail, Teams activity."],
    ) -> MyDaySummary:
        """What does my day look like?"""
        return await _my_day_v1_impl(ctx, payload)
