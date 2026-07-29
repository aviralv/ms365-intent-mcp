"""TDD tests for Task 3: start_timezone / end_timezone fields + date widening.

These tests FAIL before the schema edits (extra="forbid" rejects unknown fields,
and pure date values fail datetime coercion), and PASS after.
"""
import datetime as _dt

from ms365_intent_mcp.intent.whats_new.schemas import EventSummary as WNEvent
from ms365_intent_mcp.intent.my_day.schemas import EventSummary as MDEvent
from ms365_intent_mcp.intent.schedule.schemas import TimeSlot


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
