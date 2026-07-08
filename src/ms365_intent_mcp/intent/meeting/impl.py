"""meeting_v1 implementation — wraps composers.meeting.compose_meeting.

Structured fields (id, subject, start, end, organizer, attendees) are
placeholder stubs until Task 12 refactors composers to return
``(dict, markdown)`` tuples. ``rendered_markdown`` carries the current
composer output verbatim.
"""

from __future__ import annotations

from datetime import datetime, timezone

from fastmcp import Context

from ...composers.meeting import compose_meeting
from .._helpers import _get_deps, wrap_errors
from .schemas import MeetingDetail, MeetingPayload, PersonRef

TOOL_NAME = "meeting_v1"


@wrap_errors(TOOL_NAME)
async def _meeting_v1_impl(ctx: Context, payload: MeetingPayload) -> MeetingDetail:
    """Call the underlying composer and return a typed MeetingDetail.

    Structured fields (id, subject, start, end, organizer, attendees) are
    stubs — Task 12 fills them when the composer returns a ``(dict, markdown)``
    tuple. ``rendered_markdown`` carries the current output verbatim.
    """
    config, client, permissions = _get_deps(ctx)
    markdown = await compose_meeting(
        client,
        permissions,
        payload.identifier,
        config.default_timezone,
    )
    return MeetingDetail(
        id="pending-composer-dict-refactor",
        subject=payload.identifier,  # best guess until Task 12
        start=datetime.now(timezone.utc),  # placeholder
        end=datetime.now(timezone.utc),  # placeholder
        organizer=PersonRef(name="unknown"),
        attendees=[],
        rendered_markdown=markdown,
    )
