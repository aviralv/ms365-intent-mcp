"""FastMCP registration for compose."""

from __future__ import annotations

from typing import Annotated

from fastmcp import Context, FastMCP

from .._shared import ErrorResponse
from .impl import _compose_impl
from .schemas import ComposePayload, ComposeResponse


def register(mcp: FastMCP) -> None:
    """Register the compose tool on a FastMCP instance."""

    @mcp.tool(annotations={"destructiveHint": True, "openWorldHint": True})
    async def compose(
        ctx: Context,
        payload: Annotated[
            ComposePayload,
            "Draft or send. Pick a variant via `type`: 'email' drafts an email "
            "(new, reply, reply_all, forward), 'event' creates, forwards, or updates "
            "a calendar event. Update requires exact event_id and explicitly supplied "
            "subject/start/end/location; omitted fields are preserved. "
            "'teams_message' sends a Teams chat message. Per-mode "
            "field requirements are enforced with descriptive errors. Optional "
            "idempotency_key dedupes retries within 10 minutes.",
        ],
    ) -> ComposeResponse | ErrorResponse:
        """Draft email, create/forward/update events, or send Teams messages (typed responses)."""
        return await _compose_impl(ctx, payload)
