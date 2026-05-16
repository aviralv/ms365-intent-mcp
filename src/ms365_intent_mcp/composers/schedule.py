"""schedule composer — find meeting times via POST /me/findMeetingTimes."""

from ..formatters import format_meeting_times_markdown, format_section_error
from ..graph import GraphClient, GraphAPIError
from ..permissions import PermissionRegistry
from ._utils import _error_reason


async def compose_schedule(
    client: GraphClient,
    permissions: PermissionRegistry,
    attendees: list[dict],
    duration_minutes: int,
    constraints: dict | None,
) -> str:
    scope_msg = permissions.check("Calendars.ReadWrite")
    if scope_msg:
        return scope_msg

    attendee_list = [
        {"emailAddress": {"address": a["email"], "name": a.get("name", a["email"])}, "type": "required"}
        for a in attendees
    ]

    payload: dict = {
        "attendees": attendee_list,
        "meetingDuration": f"PT{duration_minutes}M",
        "returnSuggestionReasons": True,
        "minimumAttendeePercentage": 100,
    }

    if constraints:
        if constraints.get("start") and constraints.get("end"):
            payload["timeConstraint"] = {
                "activityDomain": "work",
                "timeslots": [
                    {
                        "start": {"dateTime": constraints["start"], "timeZone": "UTC"},
                        "end": {"dateTime": constraints["end"], "timeZone": "UTC"},
                    }
                ],
            }

    try:
        response = await client.post("/me/findMeetingTimes", payload)
    except GraphAPIError as exc:
        return format_section_error("Schedule", _error_reason(exc))

    suggestions = response.get("meetingTimeSuggestions", [])
    if not suggestions:
        reason = response.get("emptySuggestionsReason", "unknown")
        if "unavailable" in reason.lower() or reason == "AttendeesUnavailable":
            return "### Schedule\nNo available slots — all attendees are unavailable in that window."
        return f"### Schedule\nNo available meeting times found (reason: {reason})."

    return format_meeting_times_markdown(suggestions)
