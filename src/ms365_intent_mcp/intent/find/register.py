"""FastMCP registration for find_v1."""

from __future__ import annotations

from typing import Annotated

from fastmcp import Context, FastMCP

from .impl import _find_v1_impl
from .schemas import FindPayload, FindResults


def register(mcp: FastMCP) -> None:
    """Register the find_v1 tool on a FastMCP instance."""

    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def find_v1(
        ctx: Context,
        payload: Annotated[
            FindPayload,
            "Search across mail, files, Teams messages, and SharePoint pages. "
            "query is required. "
            "entity_type optionally restricts to one kind: 'email', 'file', 'message', or 'page'.",
        ],
    ) -> FindResults:
        """Search across mail, files, Teams messages, and SharePoint pages (typed responses)."""
        return await _find_v1_impl(ctx, payload)
