#!/usr/bin/env python3
"""Live end-to-end verify script for whats_new_v1.

Calls _whats_new_impl against a live Graph connection for the past
24 hours and prints the structured response fields plus rendered markdown.

Usage:
    uv run python scripts/verify_whats_new.py
    uv run python scripts/verify_whats_new.py mail
    uv run python scripts/verify_whats_new.py calendar
    uv run python scripts/verify_whats_new.py teams
"""

from __future__ import annotations

import asyncio
import datetime
import sys
from unittest.mock import MagicMock

from ms365_intent_mcp.auth import TokenManager
from ms365_intent_mcp.config import Config
from ms365_intent_mcp.graph import GraphClient
from ms365_intent_mcp.intent.whats_new.impl import _whats_new_impl
from ms365_intent_mcp.intent.whats_new.schemas import WhatsNewPayload
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


async def run(scope: str = "all") -> None:
    config = Config()
    auth = TokenManager(config)
    auth.ensure_authenticated()

    async with GraphClient(config.graph_base_url, auth.get_access_token) as client:
        permissions = PermissionRegistry.from_token_provider(auth.peek_access_token)
        ctx = _make_ctx(config, client, permissions)

        since = datetime.datetime.now(tz=datetime.timezone.utc) - datetime.timedelta(hours=24)
        since_iso = since.isoformat()

        payload = WhatsNewPayload.model_validate({"since": since_iso, "scope": scope})

        _line()
        print(f"whats_new — since={since_iso}, scope={scope!r}")
        _line()

        result = await _whats_new_impl(ctx, payload)

        _line("─")
        print(f"type:         {result.type}")
        print(f"since:        {result.since}")
        print(f"mail items:   {len(result.mail)}")
        print(f"events:       {len(result.events)}")
        print(f"teams items:  {len(result.teams)}")
        _line("─")
        print("rendered_markdown (first 500 chars):")
        print(result.rendered_markdown[:500])
        if len(result.rendered_markdown) > 500:
            print(f"... [{len(result.rendered_markdown) - 500} more chars]")
        _line()

        assert result.type == "whats_new_summary", f"Expected type='whats_new_summary', got {result.type!r}"
        print("✅ type field correct")
        print("✅ verify_whats_new PASSED")


if __name__ == "__main__":
    scope = sys.argv[1] if len(sys.argv) > 1 else "all"
    asyncio.run(run(scope))
