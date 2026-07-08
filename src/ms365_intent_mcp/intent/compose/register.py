"""FastMCP registration for compose."""

from __future__ import annotations

from typing import Annotated

from fastmcp import Context, FastMCP

from .impl import _compose_impl
from .schemas import ComposePayload, ComposeResponse


def register(mcp: FastMCP) -> None:
    """Register the compose tool on a FastMCP instance."""

    @mcp.tool(annotations={"destructiveHint": True, "openWorldHint": True})
    async def compose(
        ctx: Context,
        payload: Annotated[
            ComposePayload,
            "Create something. type='email' drafts an email — mode='new' needs to+subject; "
            "mode='reply'/'reply_all'/'forward' needs in_reply_to_message_id. "
            "type='event' creates a calendar event (timezone required). "
            "type='teams_message' sends a Teams chat message. "
            "Optional idempotency_key deduplicates retries within 10 minutes.",
        ],
    ) -> ComposeResponse:
        """Create an email draft, calendar event, or Teams message (typed responses)."""
        return await _compose_impl(ctx, payload)
