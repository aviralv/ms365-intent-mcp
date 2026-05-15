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

logging.getLogger("httpx").setLevel(logging.WARNING)

_config = Config()


@asynccontextmanager
async def lifespan(server: FastMCP):
    logger = logging.getLogger("ms365_intent_mcp")
    logger.setLevel(logging.INFO)
    if not logger.handlers:
        handler = logging.StreamHandler(sys.stderr)
        handler.setFormatter(logging.Formatter("%(asctime)s %(name)s %(levelname)s %(message)s"))
        logger.addHandler(handler)
        logger.propagate = False

    auth = TokenManager(_config)
    print("Connecting to Microsoft Graph API...", file=sys.stderr)
    auth.ensure_authenticated()
    print("Connected.", file=sys.stderr)

    permissions = PermissionRegistry.from_token(auth.get_access_token())
    cb = CircuitBreaker(
        failure_threshold=_config.cb_failure_threshold,
        recovery_timeout=_config.cb_recovery_timeout,
    )

    async with GraphClient(_config.graph_base_url, auth.get_access_token, cb=cb) as client:
        yield {
            "config": _config,
            "client": client,
            "permissions": permissions,
        }


mcp = FastMCP(
    name="ms365-intent",
    instructions=(
        "Intent-oriented Microsoft 365 MCP server.\n\n"
        "Tools:\n"
        "- my_day: What does my day look like? Calendar + mail + teams overview.\n"
        "- meeting: Tell me about this meeting. Full context by ID, subject, or 'next'.\n"
        "- compose: Create something — email drafts, reply drafts, events, Teams messages.\n\n"
        "Email drafts are saved to Drafts, never auto-sent.\n"
        "Teams messages require confirmation before sending."
    ),
    lifespan=lifespan,
)


def main():
    mcp.run()


if __name__ == "__main__":
    main()
