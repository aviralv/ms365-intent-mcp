"""schedule_v1 implementation — wraps composers.schedule.compose_schedule.

Typed list[Attendee] is adapted to the flat dict shape compose_schedule
still expects (until Task 12 refactors composers to return structured data).
"""

from __future__ import annotations

from fastmcp import Context

from ...composers.schedule import compose_schedule
from .._helpers import _get_deps, wrap_errors
from .schemas import SchedulePayload, ScheduleSuggestions

TOOL_NAME = "schedule_v1"


@wrap_errors(TOOL_NAME)
async def _schedule_v1_impl(ctx: Context, payload: SchedulePayload) -> ScheduleSuggestions:
    """Execute a schedule_v1 request via the underlying composer.

    Returns a typed ScheduleSuggestions response. On error, ``wrap_errors``
    catches ``IntentError`` / ``GraphAPIError`` and returns an ``ErrorResponse``.
    Structured TimeSlot objects are populated as Task 12 placeholder (empty list)
    until composers return (dict, markdown) pairs.
    """
    _, client, permissions = _get_deps(ctx)

    # Convert typed Attendee list to the flat dict shape composers expect.
    attendees_flat = [
        {"email": a.email, "name": a.name or a.email}
        for a in payload.attendees
    ]
    constraints_dict = None
    if payload.constraints:
        constraints_dict = payload.constraints.model_dump(exclude_none=True)

    markdown = await compose_schedule(
        client=client,
        permissions=permissions,
        attendees=attendees_flat,
        duration_minutes=payload.duration_minutes,
        constraints=constraints_dict,
    )

    return ScheduleSuggestions(
        suggestions=[],  # Task 12 fills these
        rendered_markdown=markdown,
    )
