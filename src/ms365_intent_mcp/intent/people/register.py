"""FastMCP registration for people_v1."""

from __future__ import annotations

from typing import Annotated

from fastmcp import Context, FastMCP

from .impl import _people_v1_impl
from .schemas import PeoplePayload, PersonDetail


def register(mcp: FastMCP) -> None:
    """Register the people_v1 tool on a FastMCP instance."""

    @mcp.tool()
    async def people_v1(
        ctx: Context,
        payload: Annotated[
            PeoplePayload,
            "Look up a person by name or email. Returns structured profile and recent context.",
        ],
    ) -> PersonDetail:
        """Look up a person and see recent email and Teams context (typed response)."""
        return await _people_v1_impl(ctx, payload)
