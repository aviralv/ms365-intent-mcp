"""TDD tests for Task 3: start_timezone / end_timezone fields + date widening.

These tests FAIL before the schema edits (extra="forbid" rejects unknown fields,
and pure date values fail datetime coercion), and PASS after.
"""
import datetime as _dt

from ms365_intent_mcp.intent.my_day.schemas import EventSummary as MDEvent
from ms365_intent_mcp.intent.schedule.schemas import TimeSlot
from ms365_intent_mcp.intent.whats_new.schemas import EventSummary as WNEvent


def test_whats_new_event_accepts_timezone_and_date():
    e = WNEvent(
        subject="All-day offsite",
        start=_dt.date(2026, 7, 29),
        end=_dt.date(2026, 7, 30),
        start_timezone="UTC",
        end_timezone="UTC",
    )
    assert e.start_timezone == "UTC"
    assert isinstance(e.start, _dt.date)


def test_whats_new_event_accepts_aware_datetime():
    e = WNEvent(
        subject="Standup",
        start="2026-07-29T14:00:00+02:00",
        end="2026-07-29T14:30:00+02:00",
        start_timezone="Europe/Berlin",
    )
    assert e.start.utcoffset() == _dt.timedelta(hours=2)
    assert e.end_timezone is None  # defaults to None


def test_my_day_event_accepts_timezone_fields():
    e = MDEvent(
        subject="X", start="2026-07-29T09:00:00+00:00",
        end="2026-07-29T09:30:00+00:00", start_timezone="UTC",
    )
    assert e.start_timezone == "UTC"


def test_schedule_timeslot_has_timezone_fields_but_stays_datetime():
    s = TimeSlot(
        start="2026-07-29T14:00:00+00:00", end="2026-07-29T14:30:00+00:00",
        confidence=0.9, start_timezone="UTC", end_timezone="UTC",
    )
    assert s.start_timezone == "UTC"


def test_all_day_date_string_serializes_without_time():
    """A date-only string must survive round-trip as a date string, not a datetime.

    This is the regression test for the ``datetime | date`` union ordering bug:
    Pydantic evaluates left-to-right, so with ``datetime`` first a string like
    ``"2026-07-30"`` gets coerced to ``datetime(2026,7,30,0,0)`` and serializes
    as ``"2026-07-30T00:00:00"`` — defeating the all-day passthrough.
    Reversing to ``date | datetime`` fixes it.
    """
    import json

    e = WNEvent(
        subject="Offsite",
        start="2026-07-30",
        end="2026-07-31",
        start_timezone="UTC",
        end_timezone="UTC",
    )
    dumped = json.loads(e.model_dump_json())
    assert dumped["start"] == "2026-07-30", (
        f"Expected '2026-07-30' but got {dumped['start']!r} — "
        "date-only string was coerced to datetime"
    )
    assert dumped["end"] == "2026-07-31", (
        f"Expected '2026-07-31' but got {dumped['end']!r} — "
        "date-only string was coerced to datetime"
    )


def test_timed_offset_aware_string_round_trips():
    """Reversing the union order must not break timed, offset-aware events.

    ``"2026-07-29T14:00:00+00:00"`` must fall through the ``date`` branch
    (which rejects strings with a time portion) and land on ``datetime``,
    then serialize with its UTC offset intact.
    """
    import json

    e = WNEvent(
        subject="Standup",
        start="2026-07-29T14:00:00+00:00",
        end="2026-07-29T14:30:00+00:00",
        start_timezone="UTC",
        end_timezone="UTC",
    )
    assert isinstance(e.start, _dt.datetime), (
        f"Expected datetime but got {type(e.start)} — timed string coerced to date"
    )
    dumped = json.loads(e.model_dump_json())
    # Pydantic normalizes +00:00 → Z for UTC; both are correct ISO 8601.
    # The key invariant: the serialized value must include a time component
    # (i.e. NOT be a bare date like "2026-07-29").
    assert "T" in dumped["start"], (
        f"Timed event start lost its time component: got {dumped['start']!r}"
    )
    assert "T" in dumped["end"], (
        f"Timed event end lost its time component: got {dumped['end']!r}"
    )
