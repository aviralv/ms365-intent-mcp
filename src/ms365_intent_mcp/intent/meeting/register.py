"""FastMCP registration for meeting."""

from __future__ import annotations

from typing import Annotated

from fastmcp import Context, FastMCP
from pydantic import TypeAdapter

from .._shared import ErrorResponse
from .impl import _meeting_impl
from .schemas import MeetingDetail, MeetingPayload


def register(mcp: FastMCP) -> None:
    """Register the meeting tool on a FastMCP instance."""

    # An inferred union gets wrapped in {"result": ...} by FastMCP. Declare an
    # object-root union explicitly so both outcomes keep their existing flat shape.
    output_schema = TypeAdapter(MeetingDetail | ErrorResponse).json_schema(mode="serialization")
    output_schema["type"] = "object"

    @mcp.tool(output_schema=output_schema)
    async def meeting(
        ctx: Context,
        payload: Annotated[
            MeetingPayload,
            "Full context for a meeting. identifier can be an event ID, "
            "a subject substring, or 'next' for the upcoming meeting. Optional date "
            "(YYYY-MM-DD) selects an occurrence starting that day in the configured "
            "calendar timezone, or validates an occurrence ID. Not supported with 'next'.",
        ],
    ) -> MeetingDetail | ErrorResponse:
        """Tell me about this meeting.

        Returns full context: attendees, body, Teams link, recording.
        """
        return await _meeting_impl(ctx, payload)
