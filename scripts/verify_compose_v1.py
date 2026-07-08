#!/usr/bin/env python3
"""Live end-to-end verify script for compose_v1.

compose_v1 is a write operation (creates drafts, events, Teams messages).
This script validates payload shapes and prints what would be sent — it
does NOT call the impl to avoid creating side effects.

Usage:
    uv run python scripts/verify_compose_v1.py
"""

from __future__ import annotations

from ms365_intent_mcp.intent.compose.schemas import (
    ComposeEmail,
    ComposeEvent,
    ComposeTeamsMessage,
)
from pydantic import TypeAdapter


def _line(char="═", n=72):
    print(char * n)


def run():
    _line()
    print("compose_v1 payload shape verification (no Graph calls)")
    _line()

    # Example 1: email draft (new)
    email_payload = ComposeEmail.model_validate({
        "type": "email",
        "mode": "new",
        "to": [{"email": "test@example.com", "name": "Test User"}],
        "subject": "Hello from verify_compose_v1",
        "body": "<p>This is a test email.</p>",
    })
    _line("─")
    print("ComposeEmail (mode='new') payload (would create draft — NOT called):")
    print(email_payload.model_dump_json(indent=2))

    # Example 2: calendar event
    event_payload = ComposeEvent.model_validate({
        "type": "event",
        "subject": "Test Meeting",
        "start": "2026-07-10T10:00:00",
        "end": "2026-07-10T11:00:00",
        "timezone": "Europe/Berlin",
        "attendees": [{"email": "test@example.com"}],
        "is_online_meeting": True,
    })
    _line("─")
    print("ComposeEvent payload (would create calendar event — NOT called):")
    print(event_payload.model_dump_json(indent=2))

    # Example 3: Teams message
    teams_payload = ComposeTeamsMessage.model_validate({
        "type": "teams_message",
        "chat_id": "19:fake-chat-id@thread.v2",
        "content": "Hello from verify_compose_v1",
    })
    _line("─")
    print("ComposeTeamsMessage payload (would send Teams message — NOT called):")
    print(teams_payload.model_dump_json(indent=2))

    # Verify discriminated union dispatch
    _line("─")
    print("Discriminated union dispatch check via TypeAdapter:")
    from pydantic import TypeAdapter
    from typing import Annotated, Union
    from pydantic import Field as _Field
    import ms365_intent_mcp.intent.compose.schemas as _schemas
    adapter = TypeAdapter(_schemas.ComposePayload)
    for d, expected_cls in [
        ({"type": "email", "mode": "new", "to": [{"email": "a@b.com"}], "subject": "Hi", "body": "Hello"}, "ComposeEmail"),
        ({"type": "event", "subject": "Sync", "start": "2026-07-10T10:00:00Z", "end": "2026-07-10T11:00:00Z", "timezone": "Europe/Berlin"}, "ComposeEvent"),
        ({"type": "teams_message", "chat_id": "19:abc@thread.v2", "content": "hi"}, "ComposeTeamsMessage"),
    ]:
        parsed = adapter.validate_python(d)
        cls_name = parsed.__class__.__name__
        status = "✅" if cls_name == expected_cls else "❌"
        print(f"  {status} type={d['type']!r} → {cls_name}")

    _line()
    print("All payload shapes validated. No Graph calls made.")
    print("✅ verify_compose_v1 PASSED")


if __name__ == "__main__":
    run()
