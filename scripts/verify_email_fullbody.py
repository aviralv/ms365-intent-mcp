#!/usr/bin/env python3
"""End-to-end verification for issue #19.

Drives the same two-hop flow the tool exposes:

  1. compose_find(search_type="email", query=<arg>)   → search-hit markdown
  2. resolve_url(webLink)                              → parse the follow-up URL
  3. compose_resolve(url=webLink)                      → full email body markdown

Success = the resolve output contains substantially more body content than the
find snippet — confirming the caller can go from "search" to "full body" using
only the outputs exposed by the tool.

Usage:
  uv run python scripts/verify_email_fullbody.py "Alessia PMI bug"
"""

from __future__ import annotations

import asyncio
import re
import sys

from ms365_intent_mcp.auth import TokenManager
from ms365_intent_mcp.composers.find import compose_find
from ms365_intent_mcp.composers.resolve import compose_resolve
from ms365_intent_mcp.config import Config
from ms365_intent_mcp.graph import GraphClient
from ms365_intent_mcp.permissions import PermissionRegistry


def _line(char: str = "─", n: int = 72) -> None:
    print(char * n)


async def run(query: str) -> None:
    config = Config()
    auth = TokenManager(config)
    auth.ensure_authenticated()

    async with GraphClient(config.graph_base_url, auth.get_access_token) as client:
        permissions = PermissionRegistry.from_token_provider(auth.peek_access_token)

        _line("═")
        print(f"STEP 1 — find(type='email', query={query!r})")
        _line()
        find_out = await compose_find(
            client=client,
            permissions=permissions,
            query=query,
            search_type="email",
        )
        print(find_out)
        print()

        # Extract the first webLink from find's markdown output. The formatter
        # emits it after "🔗 " on its own line.
        match = re.search(r"🔗\s+(\S+)", find_out)
        if not match:
            _line("═")
            print("❌ No 🔗 webLink found in find output — cannot proceed to resolve.")
            return
        web_link = match.group(1)

        _line("═")
        print(f"STEP 2 — resolve(url={web_link[:80]}…)")
        _line()
        resolve_out = await compose_resolve(
            client=client,
            permissions=permissions,
            url=web_link,
        )
        print(resolve_out)
        print()

        _line("═")
        print("VERIFICATION")
        _line()
        find_body_bytes = len(find_out.encode("utf-8"))
        resolve_body_bytes = len(resolve_out.encode("utf-8"))
        print(f"  find markdown:    {find_body_bytes:,} bytes")
        print(f"  resolve markdown: {resolve_body_bytes:,} bytes")
        ratio = resolve_body_bytes / max(find_body_bytes, 1)
        print(f"  ratio:            {ratio:.1f}×")
        if resolve_body_bytes > find_body_bytes:
            print("  ✅ resolve returned more content than find — full body flow works")
        else:
            print("  ⚠️  resolve did not exceed find — inspect output above")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    query = " ".join(sys.argv[1:])
    asyncio.run(run(query))
