#!/usr/bin/env python3
"""Live end-to-end verify script for meeting_v1.

Calls _meeting_impl against a live Graph connection and prints
the structured response fields plus the rendered markdown.

Usage:
    uv run python scripts/verify_meeting.py
    uv run python scripts/verify_meeting.py "next"
    uv run python scripts/verify_meeting.py "Refinement"
"""

from __future__ import annotations

import asyncio
import sys
from unittest.mock import MagicMock

from ms365_intent_mcp.auth import TokenManager
from ms365_intent_mcp.config import Config
from ms365_intent_mcp.graph import GraphClient
from ms365_intent_mcp.intent.meeting.impl import _meeting_impl
from ms365_intent_mcp.intent.meeting.schemas import MeetingPayload
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


async def run(identifier: str) -> None:
    config = Config()
    auth = TokenManager(config)
    auth.ensure_authenticated()

    async with GraphClient(config.graph_base_url, auth.get_access_token) as client:
        permissions = PermissionRegistry.from_token_provider(auth.peek_access_token)
        ctx = _make_ctx(config, client, permissions)

        payload = MeetingPayload.model_validate({"identifier": identifier})

        _line()
        print(f"meeting — identifier={identifier!r}")
        _line()

        result = await _meeting_impl(ctx, payload)

        _line("─")
        print(f"type:       {result.type}")
        print(f"id:         {result.id}")
        print(f"subject:    {result.subject}")
        print(f"start:      {result.start}")
        print(f"end:        {result.end}")
        print(f"organizer:  {result.organizer.name} <{result.organizer.email}>")
        print(f"attendees:  {len(result.attendees)}")
        print(f"recording:  {result.recording is not None}")
        _line("─")
        print("rendered_markdown (first 500 chars):")
        print(result.rendered_markdown[:500])
        if len(result.rendered_markdown) > 500:
            print(f"... [{len(result.rendered_markdown) - 500} more chars]")
        _line()

        assert result.type == "meeting_detail", f"Expected type='meeting_detail', got {result.type!r}"
        assert result.subject, "Expected non-empty subject"
        print("✅ type field correct")
        print("✅ subject non-empty")
        print("✅ verify_meeting PASSED")


if __name__ == "__main__":
    identifier = sys.argv[1] if len(sys.argv) > 1 else "next"
    asyncio.run(run(identifier))
