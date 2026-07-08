#!/usr/bin/env python3
"""Live end-to-end verify script for find_v1.

Calls _find_v1_impl against a live Graph connection and prints
the structured response fields plus the rendered markdown.

Usage:
    uv run python scripts/verify_find_v1.py "search term"
    uv run python scripts/verify_find_v1.py "budget" email
    uv run python scripts/verify_find_v1.py "report" file
"""

from __future__ import annotations

import asyncio
import sys
from unittest.mock import MagicMock

from ms365_intent_mcp.auth import TokenManager
from ms365_intent_mcp.config import Config
from ms365_intent_mcp.graph import GraphClient
from ms365_intent_mcp.intent.find.impl import _find_v1_impl
from ms365_intent_mcp.intent.find.schemas import FindPayload
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


async def run(query: str, entity_type: str | None = None) -> None:
    config = Config()
    auth = TokenManager(config)
    auth.ensure_authenticated()

    async with GraphClient(config.graph_base_url, auth.get_access_token) as client:
        permissions = PermissionRegistry.from_token_provider(auth.peek_access_token)
        ctx = _make_ctx(config, client, permissions)

        payload_dict: dict = {"query": query}
        if entity_type:
            payload_dict["entity_type"] = entity_type
        payload = FindPayload.model_validate(payload_dict)

        _line()
        print(f"find_v1 — query={query!r}, entity_type={entity_type!r}")
        _line()

        result = await _find_v1_impl(ctx, payload)

        _line("─")
        print(f"type:       {result.type}")
        print(f"query:      {result.query}")
        print(f"hits count: {len(result.hits)}")
        for i, hit in enumerate(result.hits[:5]):
            print(f"  [{i}] kind={hit.kind}")
        _line("─")
        print("rendered_markdown (first 500 chars):")
        print(result.rendered_markdown[:500])
        if len(result.rendered_markdown) > 500:
            print(f"... [{len(result.rendered_markdown) - 500} more chars]")
        _line()

        assert result.type == "find_results", f"Expected type='find_results', got {result.type!r}"
        print("✅ type field correct")
        print("✅ verify_find_v1 PASSED")


if __name__ == "__main__":
    query = sys.argv[1] if len(sys.argv) > 1 else "test"
    entity_type = sys.argv[2] if len(sys.argv) > 2 else None
    asyncio.run(run(query, entity_type))
