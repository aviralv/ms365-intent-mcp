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
        "Intent-oriented Microsoft 365 MCP. Route by intent:\n"
        "- my_day → daily overview (calendar, mail, Teams)\n"
        "- meeting → one meeting's full context\n"
        "- compose → draft email / create-or-forward event / send Teams message\n"
        "- schedule → find meeting times\n"
        "- people → look up a person\n"
        "- whats_new → what changed since a time\n"
        "- find → cross-source search\n"
        "- resolve → resolve any M365 URL\n"
        "- transcript → download a recording's VTT (or list)\n\n"
        "Email drafts save to Drafts (never auto-sent); Teams sends confirm first."
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
