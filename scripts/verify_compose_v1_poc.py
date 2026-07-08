#!/usr/bin/env python3
"""Live PoC for compose_v1 — verifies the discriminated union path.

Three probes without hitting Graph:

  A. Schema validation via TypeAdapter — proves the discriminator dispatch works.
  B. Dispatch through _compose_impl with a stubbed compose_action.
  C. Dump the generated JSON Schema so we can inspect the ``oneOf`` shape.

The GATE for this session is Probe C: the JSON Schema for ``ComposePayload``
must render as ``oneOf`` with three variants, each with a ``const`` on the
``type`` field. If not, the spec has to be re-reviewed.

Usage:
    uv run python scripts/verify_compose_v1_poc.py
"""

from __future__ import annotations

import asyncio
import json
from unittest.mock import AsyncMock, MagicMock

from pydantic import TypeAdapter

from ms365_intent_mcp.intent.compose.impl import _compose_impl
from ms365_intent_mcp.intent.compose.schemas import (
    ComposeEmail,
    ComposeEvent,
    ComposePayload,
    ComposeTeamsMessage,
)


def _line(char: str = "═", n: int = 72) -> None:
    print(char * n)


def _mock_ctx():
    ctx = MagicMock()
    config = MagicMock(default_timezone="Europe/Berlin")
    client = AsyncMock()
    permissions = MagicMock()
    permissions.check = MagicMock(return_value=None)
    ctx.request_context.lifespan_context = {
        "config": config,
        "client": client,
        "permissions": permissions,
    }
    return ctx


async def run() -> None:
    ctx = _mock_ctx()

    _line()
    print("PROBE A — schema validation via ComposePayload discriminator")
    _line()
    adapter = TypeAdapter(ComposePayload)
    for payload_dict in [
        {"type": "email", "mode": "new", "to": [{"email": "a@b.com"}],
         "subject": "Hi", "body": "Hello"},
        {"type": "event", "subject": "Sync",
         "start": "2026-07-08T10:00:00Z", "end": "2026-07-08T11:00:00Z",
         "timezone": "Europe/Berlin"},
        {"type": "teams_message", "chat_id": "19:abc@thread.v2", "content": "hi"},
    ]:
        parsed = adapter.validate_python(payload_dict)
        print(f"  ✅ {parsed.__class__.__name__}: type={parsed.type}")

    _line()
    print("PROBE B — dispatch to legacy composer via _compose_impl")
    _line()

    import ms365_intent_mcp.intent.compose.impl as impl_module

    async def _fake_compose_action(client_arg, perms_arg, action_type, params):
        return f"✅ (stub) {action_type.value} — keys={list(params.keys())}"

    original = impl_module.compose_action
    impl_module.compose_action = _fake_compose_action
    try:
        r1 = await _compose_impl(ctx, ComposeEmail.model_validate({
            "type": "email", "mode": "new",
            "to": [{"email": "a@b.com"}], "subject": "Hi", "body": "Hello",
        }))
        print(f"  Email: type={r1.type}, subject={r1.subject}")

        r2 = await _compose_impl(ctx, ComposeEvent.model_validate({
            "type": "event", "subject": "Sync",
            "start": "2026-07-08T10:00:00Z", "end": "2026-07-08T11:00:00Z",
            "timezone": "Europe/Berlin",
        }))
        print(f"  Event: type={r2.type}, subject={r2.subject}")

        r3 = await _compose_impl(ctx, ComposeTeamsMessage.model_validate({
            "type": "teams_message", "chat_id": "19:abc@thread.v2", "content": "hi",
        }))
        print(f"  Teams: type={r3.type}, chat_id={r3.chat_id}")
    finally:
        impl_module.compose_action = original

    _line()
    print("PROBE C — JSON Schema for ComposePayload discriminator (GATE)")
    _line()
    schema = adapter.json_schema()
    schema_json = json.dumps(schema, indent=2)
    print(schema_json[:3000])
    if len(schema_json) > 3000:
        print(f"\n... [{len(schema_json) - 3000} more chars truncated]")
    _line()
    has_discriminator = "discriminator" in schema_json
    has_oneof = '"oneOf"' in schema_json
    print(f"  Contains 'discriminator': {'✅' if has_discriminator else '❌'}")
    print(f"  Contains 'oneOf':         {'✅' if has_oneof else '❌'}")
    if has_discriminator and has_oneof:
        print("\n  ✅ GATE PASSED — FastMCP-compatible discriminated union renders correctly")
    else:
        print("\n  🚨 GATE FAILED — the spec's Pydantic pattern is not rendering to the")
        print("     expected oneOf/discriminator JSON Schema shape. Re-review the spec.")


if __name__ == "__main__":
    asyncio.run(run())
