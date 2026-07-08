"""FastMCP registration for people."""

from __future__ import annotations

from typing import Annotated

from fastmcp import Context, FastMCP

from .impl import _people_impl
from .schemas import PeoplePayload, PersonDetail


def register(mcp: FastMCP) -> None:
    """Register the people tool on a FastMCP instance."""

    @mcp.tool()
    async def people(
        ctx: Context,
        payload: Annotated[
            PeoplePayload,
            "Look up a person by name or email. Returns structured profile and recent context.",
        ],
    ) -> PersonDetail:
        """Look up a person and see recent email and Teams context (typed response)."""
        return await _people_impl(ctx, payload)
