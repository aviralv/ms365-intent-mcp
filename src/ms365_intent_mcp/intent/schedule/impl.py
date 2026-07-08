"""schedule_v1 implementation — wraps composers.schedule.compose_schedule."""

from __future__ import annotations

from fastmcp import Context

from ...composers.schedule import compose_schedule
from .._helpers import _get_deps, wrap_errors
from .schemas import SchedulePayload, ScheduleSuggestions, TimeSlot

TOOL_NAME = "schedule_v1"


@wrap_errors(TOOL_NAME)
async def _schedule_v1_impl(ctx: Context, payload: SchedulePayload) -> ScheduleSuggestions:
    """Execute a schedule_v1 request via the underlying composer."""
    _, client, permissions = _get_deps(ctx)

    attendees_flat = [
        {"email": a.email, "name": a.name or a.email}
        for a in payload.attendees
    ]
    constraints_dict = None
    if payload.constraints:
        constraints_dict = payload.constraints.model_dump(exclude_none=True)

    data, markdown = await compose_schedule(
        client=client,
        permissions=permissions,
        attendees=attendees_flat,
        duration_minutes=payload.duration_minutes,
        constraints=constraints_dict,
    )

    suggestions = []
    for s in data.get("suggestions", []):
        try:
            suggestions.append(TimeSlot.model_validate(s))
        except Exception:
            pass

    return ScheduleSuggestions(
        suggestions=suggestions,
        rendered_markdown=markdown,
    )
