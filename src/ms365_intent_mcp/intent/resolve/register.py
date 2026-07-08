"""FastMCP registration for resolve."""

from __future__ import annotations

from typing import Annotated

from fastmcp import Context, FastMCP

from .impl import _resolve_impl
from .schemas import ResolvePayload, ResolvedContent


def register(mcp: FastMCP) -> None:
    """Register the resolve tool on a FastMCP instance."""

    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def resolve(
        ctx: Context,
        payload: Annotated[
            ResolvePayload,
            "Resolve any M365 URL and return its content. "
            "url: any Teams message/meeting link, Outlook deep link, "
            "SharePoint page URL, or OneDrive file URL. "
            "The URL type is auto-detected — no discriminator needed.",
        ],
    ) -> ResolvedContent:
        """Resolve any Microsoft 365 URL and return its content (typed response)."""
        return await _resolve_impl(ctx, payload)
