"""Schema-level tests for the include_bodies additions."""

from ms365_intent_mcp.intent.my_day.schemas import EventSummary, MyDayPayload


def test_payload_include_bodies_defaults_false():
    assert MyDayPayload().include_bodies is False


def test_payload_accepts_include_bodies_true():
    assert MyDayPayload(include_bodies=True).include_bodies is True


def test_event_summary_body_links_default():
    ev = EventSummary(subject="x", start="2026-08-05", end="2026-08-05")
    assert ev.body is None
    assert ev.links == []


def test_event_summary_accepts_body_links():
    ev = EventSummary(
        subject="x",
        start="2026-08-05",
        end="2026-08-05",
        body="agenda",
        links=["https://a.example.com"],
    )
    assert ev.body == "agenda"
    assert ev.links == ["https://a.example.com"]
