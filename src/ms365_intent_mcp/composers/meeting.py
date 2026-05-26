"""meeting composer — full context for a single meeting."""

from datetime import datetime, timedelta

from ..formatters import format_event_detail_markdown
from ..graph import GraphClient, GraphAPIError
from ..permissions import PermissionRegistry
from ._utils import _escape_odata


async def compose_meeting(
    client: GraphClient,
    permissions: PermissionRegistry,
    identifier: str,
    timezone: str,
) -> str:
    event = await _resolve_event(client, identifier, timezone)
    if event is None:
        return f'❌ No meeting found matching "{identifier}"'

    return format_event_detail_markdown(event)


async def _resolve_event(
    client: GraphClient,
    identifier: str,
    timezone: str,
) -> dict | None:
    if identifier.lower() == "next":
        return await _find_next_event(client, timezone)

    if " " not in identifier:
        try:
            headers = GraphClient.calendar_headers(timezone)
            return await client.get(f"/me/events/{identifier}", headers=headers)
        except GraphAPIError:
            pass

    return await _search_by_subject(client, identifier, timezone)


async def _find_next_event(client: GraphClient, timezone: str) -> dict | None:
    now = datetime.now()
    start_iso = now.strftime("%Y-%m-%dT%H:%M:%S")
    end_iso = (now + timedelta(days=1)).strftime("%Y-%m-%dT23:59:59")

    headers = GraphClient.calendar_headers(timezone)
    result = await client.get("/me/calendarView", params={
        "startDateTime": start_iso,
        "endDateTime": end_iso,
        "$orderby": "start/dateTime",
        "$top": "1",
    }, headers=headers)

    events = result.get("value", [])
    return events[0] if events else None


async def _search_by_subject(
    client: GraphClient,
    subject: str,
    timezone: str,
) -> dict | None:
    now = datetime.now()
    start_iso = (now - timedelta(days=7)).strftime("%Y-%m-%dT00:00:00")
    end_iso = (now + timedelta(days=7)).strftime("%Y-%m-%dT23:59:59")

    headers = GraphClient.calendar_headers(timezone)

    try:
        result = await client.get("/me/calendarView", params={
            "startDateTime": start_iso,
            "endDateTime": end_iso,
            "$orderby": "start/dateTime asc",
            "$top": "50",
            "$filter": f"contains(subject, '{_escape_odata(subject)}')",
        }, headers=headers)

        events = result.get("value", [])
        if events:
            return events[0]
    except GraphAPIError:
        pass

    result = await client.get("/me/calendarView", params={
        "startDateTime": start_iso,
        "endDateTime": end_iso,
        "$orderby": "start/dateTime asc",
        "$top": "50",
    }, headers=headers)

    subject_lower = subject.lower()
    for event in result.get("value", []):
        if subject_lower in event.get("subject", "").lower():
            return event

    return None
