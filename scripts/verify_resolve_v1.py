#!/usr/bin/env python3
"""Live end-to-end verify script for resolve_v1.

Calls _resolve_v1_impl against a live Graph connection and prints
the structured response fields plus the rendered markdown.

Usage:
    uv run python scripts/verify_resolve_v1.py <M365-URL>

Examples:
    uv run python scripts/verify_resolve_v1.py "https://teams.microsoft.com/l/chat/19:xxx@thread.v2"
    uv run python scripts/verify_resolve_v1.py "https://outlook.office.com/mail/id/AAA..."
"""

from __future__ import annotations

import asyncio
import sys
from unittest.mock import MagicMock

from ms365_intent_mcp.auth import TokenManager
from ms365_intent_mcp.config import Config
from ms365_intent_mcp.graph import GraphClient
from ms365_intent_mcp.intent.resolve.impl import _resolve_v1_impl
from ms365_intent_mcp.intent.resolve.schemas import ResolvePayload
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


async def run(url: str) -> None:
    config = Config()
    auth = TokenManager(config)
    auth.ensure_authenticated()

    async with GraphClient(config.graph_base_url, auth.get_access_token) as client:
        permissions = PermissionRegistry.from_token_provider(auth.peek_access_token)
        ctx = _make_ctx(config, client, permissions)

        payload = ResolvePayload.model_validate({"url": url})

        _line()
        print(f"resolve_v1 — url={url[:80]}{'…' if len(url) > 80 else ''}")
        _line()

        result = await _resolve_v1_impl(ctx, payload)

        _line("─")
        print(f"type:     {result.type}")
        print(f"kind:     {result.kind}")
        print(f"url:      {result.url}")
        print(f"data cls: {result.data.__class__.__name__}")
        _line("─")
        print("rendered_markdown (first 500 chars):")
        print(result.rendered_markdown[:500])
        if len(result.rendered_markdown) > 500:
            print(f"... [{len(result.rendered_markdown) - 500} more chars]")
        _line()

        assert result.type == "resolved_content", f"Expected type='resolved_content', got {result.type!r}"
        print("✅ type field correct")
        print("✅ verify_resolve_v1 PASSED")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    asyncio.run(run(sys.argv[1]))
