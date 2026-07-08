#!/usr/bin/env python3
"""Live end-to-end verify script for my_day_v1.

Calls _my_day_v1_impl against a live Graph connection and prints
the structured response fields plus the rendered markdown.

Usage:
    uv run python scripts/verify_my_day_v1.py
    uv run python scripts/verify_my_day_v1.py 2026-07-08
"""

from __future__ import annotations

import asyncio
import sys
from datetime import date
from unittest.mock import MagicMock

from ms365_intent_mcp.auth import TokenManager
from ms365_intent_mcp.config import Config
from ms365_intent_mcp.graph import GraphClient
from ms365_intent_mcp.intent.my_day.impl import _my_day_v1_impl
from ms365_intent_mcp.intent.my_day.schemas import MyDayPayload
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


async def run(target_date: date | None = None) -> None:
    config = Config()
    auth = TokenManager(config)
    auth.ensure_authenticated()

    async with GraphClient(config.graph_base_url, auth.get_access_token) as client:
        permissions = PermissionRegistry.from_token_provider(auth.peek_access_token)
        ctx = _make_ctx(config, client, permissions)

        payload = MyDayPayload.model_validate(
            {"date": target_date.isoformat()} if target_date else {}
        )

        _line()
        print(f"my_day_v1 — date={payload.date or 'today (server default)'}")
        _line()

        result = await _my_day_v1_impl(ctx, payload)

        _line("─")
        print(f"type:           {result.type}")
        print(f"date:           {result.date}")
        print(f"events count:   {len(result.events)}")
        print(f"mail unread:    {result.mail.unread_count}")
        print(f"teams messages: {result.teams.recent_message_count}")
        _line("─")
        print("rendered_markdown (first 500 chars):")
        print(result.rendered_markdown[:500])
        if len(result.rendered_markdown) > 500:
            print(f"... [{len(result.rendered_markdown) - 500} more chars]")
        _line()

        assert result.type == "my_day_summary", f"Expected type='my_day_summary', got {result.type!r}"
        print("✅ type field correct")
        print("✅ verify_my_day_v1 PASSED")


if __name__ == "__main__":
    target_date = None
    if len(sys.argv) > 1:
        target_date = date.fromisoformat(sys.argv[1])
    asyncio.run(run(target_date))
