"""FastMCP server with lifespan managing auth, graph client, and permissions."""

import logging
import sys
from contextlib import asynccontextmanager

from fastmcp import FastMCP

from .auth import TokenManager
from .config import Config
from .graph import GraphClient
from .permissions import PermissionRegistry
from .resilience import CircuitBreaker
from .vroom import VroomClient

logging.getLogger("httpx").setLevel(logging.WARNING)


@asynccontextmanager
async def lifespan(server: FastMCP):
    logger = logging.getLogger("ms365_intent_mcp")
    logger.setLevel(logging.INFO)
    if not logger.handlers:
        handler = logging.StreamHandler(sys.stderr)
        handler.setFormatter(logging.Formatter("%(asctime)s %(name)s %(levelname)s %(message)s"))
        logger.addHandler(handler)
        logger.propagate = False

    config = Config()
    auth = TokenManager(config)
    print("Connecting to Microsoft Graph API...", file=sys.stderr)
    auth.ensure_authenticated()
    print("Connected.", file=sys.stderr)

    permissions = PermissionRegistry.from_token_provider(auth.peek_access_token)
    cb = CircuitBreaker(
        failure_threshold=config.cb_failure_threshold,
        recovery_timeout=config.cb_recovery_timeout,
    )

    async with GraphClient(config.graph_base_url, auth.get_access_token, cb=cb) as client:
        async with VroomClient(auth.get_sharepoint_token) as vroom:
            yield {
                "config": config,
                "client": client,
                "permissions": permissions,
                "vroom": vroom,
            }


mcp = FastMCP(
    name="ms365-intent",
    instructions=(
        "Intent-oriented Microsoft 365 MCP server.\n\n"
        "Tools:\n"
        "- my_day: Daily overview (payload={date?}) — calendar, mail, Teams.\n"
        "- meeting: Full context for a meeting (payload={identifier}) — event ID, subject, or 'next'.\n"
        "- compose: Create drafts (payload={type: email|event|teams_message, ...}) — discriminated union.\n"
        "- schedule: Find meeting times (payload={attendees, duration_minutes?, ...}).\n"
        "- people: Look up a person (payload={query}).\n"
        "- whats_new: What changed since X (payload={since, scope?}).\n"
        "- find: Cross-source search (payload={query, entity_type?}).\n"
        "- resolve: Resolve any M365 URL (payload={url}).\n"
        "- transcript: Download a meeting recording's VTT "
        "(payload={url?|name?|item_id+drive_id+site_root?|list?, output_dir?}).\n\n"
        "Email drafts are saved to Drafts, never auto-sent.\n"
        "Teams messages require confirmation before sending."
    ),
    lifespan=lifespan,
)

from .intent import register_all as _register_intent_surface
_register_intent_surface(mcp)


def main():
    # Subcommand dispatch before FastMCP startup — `auth` needs to bootstrap
    # a token before the lifespan (which calls ensure_authenticated) can run.
    if len(sys.argv) > 1 and sys.argv[1] == "auth":
        from .auth_cli import main as auth_main

        sys.exit(auth_main(sys.argv[2:]))
    mcp.run()


if __name__ == "__main__":
    main()
