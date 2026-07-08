#!/usr/bin/env python3
"""Live end-to-end verify script for people_v1.

Calls _people_impl against a live Graph connection and prints
the structured response fields plus the rendered markdown.

Usage:
    uv run python scripts/verify_people.py "Firstname Lastname"
    uv run python scripts/verify_people.py "firstname.lastname@example.com"
"""

from __future__ import annotations

import asyncio
import sys
from unittest.mock import MagicMock

from ms365_intent_mcp.auth import TokenManager
from ms365_intent_mcp.config import Config
from ms365_intent_mcp.graph import GraphClient
from ms365_intent_mcp.intent.people.impl import _people_impl
from ms365_intent_mcp.intent.people.schemas import PeoplePayload
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


async def run(query: str) -> None:
    config = Config()
    auth = TokenManager(config)
    auth.ensure_authenticated()

    async with GraphClient(config.graph_base_url, auth.get_access_token) as client:
        permissions = PermissionRegistry.from_token_provider(auth.peek_access_token)
        ctx = _make_ctx(config, client, permissions)

        payload = PeoplePayload.model_validate({"query": query})

        _line()
        print(f"people — query={query!r}")
        _line()

        result = await _people_impl(ctx, payload)

        _line("─")
        print(f"type:           {result.type}")
        print(f"name:           {result.name}")
        print(f"email:          {result.email}")
        print(f"job_title:      {result.job_title}")
        print(f"recent_mail:    {len(result.recent_mail)} items")
        print(f"recent_chat:    {result.recent_chat is not None}")
        _line("─")
        print("rendered_markdown (first 500 chars):")
        print(result.rendered_markdown[:500])
        if len(result.rendered_markdown) > 500:
            print(f"... [{len(result.rendered_markdown) - 500} more chars]")
        _line()

        assert result.type == "person_detail", f"Expected type='person_detail', got {result.type!r}"
        print("✅ type field correct")
        print("✅ verify_people PASSED")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    query = " ".join(sys.argv[1:])
    asyncio.run(run(query))
