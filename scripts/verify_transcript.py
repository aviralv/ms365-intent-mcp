#!/usr/bin/env python3
"""Live end-to-end verify script for transcript.

Calls _transcript_impl against a live Graph + Vroom connection and prints
the structured response plus rendered markdown.

Usage:
    uv run python scripts/verify_transcript.py --name "Sprint Review"
    uv run python scripts/verify_transcript.py --url "<recording-url-or-vroom_url>"
    uv run python scripts/verify_transcript.py --name "Fact Sheet" --out /tmp/vtt
"""

from __future__ import annotations

import argparse
import asyncio
from unittest.mock import MagicMock

from ms365_intent_mcp.auth import TokenManager
from ms365_intent_mcp.config import Config
from ms365_intent_mcp.graph import GraphClient
from ms365_intent_mcp.intent.transcript.impl import _transcript_impl
from ms365_intent_mcp.intent.transcript.schemas import TranscriptPayload
from ms365_intent_mcp.permissions import PermissionRegistry
from ms365_intent_mcp.vroom import VroomClient


def _line(char: str = "═", n: int = 72) -> None:
    print(char * n)


def _make_ctx(config, client, permissions, vroom):
    ctx = MagicMock()
    ctx.request_context.lifespan_context = {
        "config": config,
        "client": client,
        "permissions": permissions,
        "vroom": vroom,
    }
    return ctx


async def run(url: str | None, name: str | None, out: str | None) -> None:
    config = Config()
    auth = TokenManager(config)
    auth.ensure_authenticated()

    async with GraphClient(config.graph_base_url, auth.get_access_token) as client:
        async with VroomClient(auth.get_sharepoint_token) as vroom:
            permissions = PermissionRegistry.from_token_provider(auth.peek_access_token)
            ctx = _make_ctx(config, client, permissions, vroom)

            payload = TranscriptPayload.model_validate(
                {"url": url, "name": name, "output_dir": out}
            )

            _line()
            print(f"transcript — url={url!r} name={name!r}")
            _line()

            result = await _transcript_impl(ctx, payload)

            _line("─")
            print(f"type:          {result.type}")
            print(f"status:        {getattr(result, 'status', '<n/a>')}")
            print(f"file_path:     {getattr(result, 'file_path', None)}")
            print(f"meeting_name:  {getattr(result, 'meeting_name', None)}")
            print(f"meeting_date:  {getattr(result, 'meeting_date', None)}")
            print(f"line_count:    {getattr(result, 'line_count', None)}")
            print(f"byte_count:    {getattr(result, 'byte_count', None)}")
            print(f"speaker_tags:  {getattr(result, 'has_speaker_tags', None)}")
            print(f"message:       {getattr(result, 'message', None)}")
            _line("─")
            print("rendered_markdown:")
            print(result.rendered_markdown)
            _line()


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--url", default=None)
    parser.add_argument("--name", default=None)
    parser.add_argument("--out", default=None)
    args = parser.parse_args()
    if not (args.url or args.name):
        parser.error("provide --url or --name")
    asyncio.run(run(args.url, args.name, args.out))
