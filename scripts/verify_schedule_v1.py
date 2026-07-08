#!/usr/bin/env python3
"""Live end-to-end verify script for schedule_v1.

Calls _schedule_v1_impl against a live Graph connection and prints
the structured response fields plus the rendered markdown.

Usage:
    uv run python scripts/verify_schedule_v1.py
    uv run python scripts/verify_schedule_v1.py attendee@example.com
"""

from __future__ import annotations

import asyncio
import sys
from unittest.mock import MagicMock

from ms365_intent_mcp.auth import TokenManager
from ms365_intent_mcp.config import Config
from ms365_intent_mcp.graph import GraphClient
from ms365_intent_mcp.intent.schedule.impl import _schedule_v1_impl
from ms365_intent_mcp.intent.schedule.schemas import SchedulePayload
from ms365_intent_mcp.permissions import PermissionRegistry


def _line(char: str = "═", n: int = 72) -> None:
    print(char * n)


def _make_ctx(config, client, permissions):
    ctx = MagicMock()
    ctx.request_context.lifespan_context = {
        "config": config,
        "client": client,
        "permissions": permissions,
    }
    return ctx


async def run(attendee_email: str) -> None:
    config = Config()
    auth = TokenManager(config)
    auth.ensure_authenticated()

    async with GraphClient(config.graph_base_url, auth.get_access_token) as client:
        permissions = PermissionRegistry.from_token_provider(auth.peek_access_token)
        ctx = _make_ctx(config, client, permissions)

        payload = SchedulePayload.model_validate({
            "attendees": [{"email": attendee_email}],
            "duration_minutes": 30,
        })

        _line()
        print(f"schedule_v1 — attendees=[{attendee_email}], duration=30min")
        _line()

        result = await _schedule_v1_impl(ctx, payload)

        _line("─")
        print(f"type:              {result.type}")
        print(f"suggestions count: {len(result.suggestions)}")
        for i, s in enumerate(result.suggestions[:3]):
            print(f"  [{i}] {s.start} → {s.end}  confidence={s.confidence:.2f}")
        _line("─")
        print("rendered_markdown (first 500 chars):")
        print(result.rendered_markdown[:500])
        if len(result.rendered_markdown) > 500:
            print(f"... [{len(result.rendered_markdown) - 500} more chars]")
        _line()

        assert result.type == "schedule_suggestions", f"Expected type='schedule_suggestions', got {result.type!r}"
        print("✅ type field correct")
        print("✅ verify_schedule_v1 PASSED")


if __name__ == "__main__":
    attendee_email = sys.argv[1] if len(sys.argv) > 1 else "me@example.com"
    asyncio.run(run(attendee_email))
