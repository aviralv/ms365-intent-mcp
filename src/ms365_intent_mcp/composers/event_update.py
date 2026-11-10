"""Targeted calendar PATCH: preserve every field not explicitly changed."""

from datetime import UTC, datetime
from urllib.parse import quote
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from ..formatters import format_event_updated_markdown
from ..graph import GraphClient
from ..intent._helpers import IntentError


def _update_datetime(value: str, timezone: str | None) -> datetime:
    """Resolve wall clocks strictly; explicit offsets disambiguate DST folds."""
    try:
        zone = ZoneInfo(timezone or "UTC")
        dt = datetime.fromisoformat(value)
    except (ValueError, ZoneInfoNotFoundError) as exc:
        raise IntentError("validation_error", "Invalid datetime or IANA timezone") from exc
    if dt.tzinfo is not None:
        return dt.astimezone(zone)
    if timezone is None:
        raise IntentError("validation_error", "Naive timestamps require an IANA timezone")
    first = dt.replace(tzinfo=zone, fold=0)
    second = dt.replace(tzinfo=zone, fold=1)
    if first.astimezone(UTC).astimezone(zone).replace(tzinfo=None) != dt:
        raise IntentError("validation_error", "Nonexistent local time during DST transition")
    if first.utcoffset() != second.utcoffset():
        raise IntentError("validation_error", "Ambiguous local time; supply an explicit UTC offset")
    return first


def _stored_datetime(event: dict, field: str) -> datetime:
    pair = event.get(field) or {}
    # The GET explicitly requests UTC; fail closed if Graph cannot supply it.
    if pair.get("timeZone") != "UTC" or not pair.get("dateTime"):
        raise IntentError("graph_api_error", "Graph did not return event timing in requested UTC")
    try:
        return _update_datetime(pair["dateTime"], "UTC")
    except IntentError as exc:
        raise IntentError("graph_api_error", "Graph returned invalid event timing") from exc


async def update_event(client: GraphClient, params: dict) -> tuple[dict, str]:
    event_id = params.get("event_id")
    changes = {key: params[key] for key in ("subject", "start", "end", "location") if key in params}
    if not event_id or not event_id.strip() or not changes:
        raise IntentError(
            "validation_error", "update requires an exact event_id and changed fields"
        )
    if any(value is None for value in changes.values()):
        raise IntentError("validation_error", "update fields cannot be null")
    timezone = params.get("timezone") or "UTC"
    try:
        zone = ZoneInfo(timezone)
    except (ValueError, ZoneInfoNotFoundError) as exc:
        raise IntentError("validation_error", "Invalid IANA timezone") from exc
    times = {
        field: _update_datetime(changes[field], params.get("timezone"))
        for field in ("start", "end")
        if field in changes
    }
    endpoint = f"/me/events/{quote(event_id, safe='')}"
    headers = GraphClient.calendar_headers("UTC")
    current = await client.get(endpoint, headers=headers)
    if current.get("type") == "seriesMaster":
        raise IntentError("validation_error", "Use an occurrence ID, not a recurring series master")
    if current.get("isAllDay") and changes.keys() & {"start", "end"}:
        raise IntentError("validation_error", "Time updates of all-day events are not supported")
    start = _stored_datetime(current, "start")
    end = _stored_datetime(current, "end")
    patch: dict = {}
    for field, dt in times.items():
        stored = start if field == "start" else end
        if dt.astimezone(UTC) != stored.astimezone(UTC):
            patch[field] = {
                "dateTime": dt.astimezone(UTC).replace(tzinfo=None).isoformat(),
                "timeZone": "UTC",
            }
        if field == "start":
            start = dt
        else:
            end = dt
    if changes.keys() & {"start", "end"} and end.astimezone(UTC) <= start.astimezone(UTC):
        raise IntentError("validation_error", "end must be after start")
    if "subject" in changes:
        if not changes["subject"].strip():
            raise IntentError("validation_error", "updated subject cannot be empty")
        if changes["subject"] != current.get("subject"):
            patch["subject"] = changes["subject"]
    if "location" in changes:
        if changes["location"] != (current.get("location") or {}).get("displayName", ""):
            patch["location"] = {"displayName": changes["location"]}
    if not patch:
        raise IntentError("validation_error", "No event fields would change")
    # An ETag, when supplied by Graph, prevents stale-read updates. A 412 is
    # surfaced by the existing Graph error wrapper; do not retry unconditionally.
    if current.get("@odata.etag"):
        headers = {**headers, "If-Match": current["@odata.etag"]}
    event = await client.patch(endpoint, patch, headers=headers)
    if not event or event.get("id") != current.get("id"):
        raise IntentError("graph_api_error", "Graph PATCH did not return the updated event")
    start = _stored_datetime(event, "start").astimezone(zone)
    end = _stored_datetime(event, "end").astimezone(zone)
    data = {
        "event_id": event["id"],
        "subject": event.get("subject", "(no subject)"),
        "start": start.isoformat(),
        "end": end.isoformat(),
        "timezone": timezone,
        "web_link": event.get("webLink") or None,
    }
    display = {
        **event,
        "start": {"dateTime": start.isoformat(), "timeZone": timezone},
        "end": {"dateTime": end.isoformat(), "timeZone": timezone},
    }
    return data, format_event_updated_markdown(display)
