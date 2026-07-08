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
) -> tuple[dict, str]:
    scope_msg = permissions.check("Calendars.ReadWrite")
    if scope_msg:
        return {"suggestions": []}, scope_msg

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
        markdown = format_section_error("Schedule", _error_reason(exc))
        return {"suggestions": []}, markdown

    suggestions = response.get("meetingTimeSuggestions", [])
    if not suggestions:
        reason = response.get("emptySuggestionsReason", "unknown")
        if "unavailable" in reason.lower() or reason == "AttendeesUnavailable":
            markdown = "### Schedule\nNo available slots — all attendees are unavailable in that window."
        else:
            markdown = f"### Schedule\nNo available meeting times found (reason: {reason})."
        return {"suggestions": []}, markdown

    markdown = format_meeting_times_markdown(suggestions)
    data_suggestions = []
    for s in suggestions:
        slot = s.get("meetingTimeSlot", {})
        start_dt = slot.get("start", {}).get("dateTime", "")
        end_dt = slot.get("end", {}).get("dateTime", "")
        start_tz = slot.get("start", {}).get("timeZone", "UTC")
        confidence = s.get("confidence", 0.0) / 100.0  # Graph gives 0-100, TimeSlot wants 0.0-1.0
        if start_dt and end_dt:
            data_suggestions.append({
                "start": start_dt if start_tz == "UTC" else f"{start_dt}",
                "end": end_dt,
                "confidence": confidence,
            })
    return {"suggestions": data_suggestions}, markdown
