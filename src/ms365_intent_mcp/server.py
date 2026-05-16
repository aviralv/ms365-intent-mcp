"""FastMCP server with lifespan managing auth, graph client, and permissions."""

import logging
import sys
from contextlib import asynccontextmanager
from datetime import datetime
from typing import Annotated

from fastmcp import Context, FastMCP

from .auth import TokenManager
from .composers.compose import ComposeType, compose_action
from .composers.find import compose_find
from .composers.meeting import compose_meeting
from .composers.my_day import compose_my_day
from .composers.people import compose_people
from .composers.resolve import compose_resolve
from .composers.schedule import compose_schedule
from .composers.whats_new import compose_whats_new
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
        "- my_day: Daily overview — calendar, mail, Teams.\n"
        "- meeting: Full context for a meeting by ID, subject, or 'next'.\n"
        "- compose: Create email drafts, reply drafts, calendar events, Teams messages.\n"
        "- schedule: Find available meeting times for a set of attendees.\n"
        "- people: Look up a person and see recent context.\n"
        "- whats_new: What happened since a given time across mail/calendar/Teams.\n"
        "- find: Search across mail, files, Teams messages, and SharePoint.\n"
        "- resolve: Resolve any M365 URL (Teams link, Outlook deep link, SharePoint page, OneDrive file).\n\n"
        "Email drafts are saved to Drafts, never auto-sent.\n"
        "Teams messages require confirmation before sending."
    ),
    lifespan=lifespan,
)


@mcp.tool()
async def my_day(
    ctx: Context,
    date: Annotated[str | None, "Date to show (YYYY-MM-DD). Defaults to today."] = None,
) -> str:
    """What does my day look like? Returns calendar events, mail summary, and Teams activity."""
    config: Config = ctx.request_context.lifespan_context["config"]
    client: GraphClient = ctx.request_context.lifespan_context["client"]
    permissions: PermissionRegistry = ctx.request_context.lifespan_context["permissions"]

    target_date = date or datetime.now().strftime("%Y-%m-%d")
    return await compose_my_day(
        client=client,
        permissions=permissions,
        date=target_date,
        timezone=config.default_timezone,
    )


@mcp.tool()
async def meeting(
    ctx: Context,
    identifier: Annotated[str, "Event ID, subject text to search, or 'next' for upcoming"],
) -> str:
    """Tell me about this meeting. Returns full context: attendees, body, Teams link."""
    config: Config = ctx.request_context.lifespan_context["config"]
    client: GraphClient = ctx.request_context.lifespan_context["client"]
    permissions: PermissionRegistry = ctx.request_context.lifespan_context["permissions"]

    return await compose_meeting(
        client=client,
        permissions=permissions,
        identifier=identifier,
        timezone=config.default_timezone,
    )


@mcp.tool()
async def compose(
    ctx: Context,
    type: Annotated[ComposeType, "What to create: email_draft, reply_draft, event, teams_message"],
    subject: Annotated[str | None, "Subject/title (for email_draft and event)"] = None,
    body: Annotated[str | None, "Body content (HTML supported for emails)"] = None,
    to: Annotated[list[dict] | None, "Recipients: [{'email': '...', 'name': '...'}]"] = None,
    cc: Annotated[list[dict] | None, "CC recipients (email_draft only)"] = None,
    message_id: Annotated[str | None, "Message ID to reply to (reply_draft only)"] = None,
    reply_all: Annotated[bool, "Reply to all (reply_draft only)"] = True,
    start: Annotated[str | None, "Start datetime YYYY-MM-DDTHH:MM:SS (event only)"] = None,
    end: Annotated[str | None, "End datetime YYYY-MM-DDTHH:MM:SS (event only)"] = None,
    attendees: Annotated[list[dict] | None, "Attendees: [{'email': '...'}] (event only)"] = None,
    location: Annotated[str | None, "Location (event only)"] = None,
    is_online_meeting: Annotated[bool, "Create as Teams meeting (event only)"] = False,
    chat_id: Annotated[str | None, "Teams chat ID (teams_message only)"] = None,
    content: Annotated[str | None, "Message content (teams_message only)"] = None,
) -> str:
    """Create something: email draft, reply draft, calendar event, or Teams message."""
    config: Config = ctx.request_context.lifespan_context["config"]
    client: GraphClient = ctx.request_context.lifespan_context["client"]
    permissions: PermissionRegistry = ctx.request_context.lifespan_context["permissions"]

    params = {
        "subject": subject,
        "body": body,
        "to": to,
        "cc": cc,
        "message_id": message_id,
        "reply_all": reply_all,
        "start": start,
        "end": end,
        "attendees": attendees,
        "location": location,
        "is_online_meeting": is_online_meeting,
        "chat_id": chat_id,
        "content": content,
        "timezone": config.default_timezone,
        "importance": "normal",
    }
    params = {k: v for k, v in params.items() if v is not None}

    return await compose_action(
        client=client,
        permissions=permissions,
        action_type=type,
        params=params,
    )


@mcp.tool()
async def schedule(
    ctx: Context,
    attendees: Annotated[list[dict], "Attendees: [{'email': '...', 'name': '...'}]"],
    duration_minutes: Annotated[int, "Meeting duration in minutes (default 30)"] = 30,
    constraints: Annotated[dict | None, "Optional time constraints: {'start': 'ISO', 'end': 'ISO'}"] = None,
) -> str:
    """Find available meeting times. Returns ranked time slots with confidence scores."""
    client: GraphClient = ctx.request_context.lifespan_context["client"]
    permissions: PermissionRegistry = ctx.request_context.lifespan_context["permissions"]
    return await compose_schedule(
        client=client,
        permissions=permissions,
        attendees=attendees,
        duration_minutes=duration_minutes,
        constraints=constraints,
    )


@mcp.tool()
async def people(
    ctx: Context,
    query: Annotated[str, "Name or email to search for"],
) -> str:
    """Look up a person and see recent email and Teams context."""
    client: GraphClient = ctx.request_context.lifespan_context["client"]
    permissions: PermissionRegistry = ctx.request_context.lifespan_context["permissions"]
    return await compose_people(client=client, permissions=permissions, query=query)


@mcp.tool()
async def whats_new(
    ctx: Context,
    since: Annotated[str, "ISO datetime to look back from, e.g. '2026-05-14T00:00:00'"],
    scope: Annotated[str | None, "Filter to: 'mail', 'calendar', 'teams', or 'all' (default)"] = None,
) -> str:
    """What happened since a given time? Returns new mail, events, and Teams messages."""
    config: Config = ctx.request_context.lifespan_context["config"]
    client: GraphClient = ctx.request_context.lifespan_context["client"]
    permissions: PermissionRegistry = ctx.request_context.lifespan_context["permissions"]
    return await compose_whats_new(
        client=client,
        permissions=permissions,
        since=since,
        scope=scope,
        timezone=config.default_timezone,
    )


@mcp.tool()
async def find(
    ctx: Context,
    query: Annotated[str, "Search query text"],
    type: Annotated[str | None, "Optional filter: 'email', 'file', 'message' (Teams), 'page' (SharePoint)"] = None,
) -> str:
    """Search across mail, files, Teams messages, and SharePoint pages."""
    client: GraphClient = ctx.request_context.lifespan_context["client"]
    permissions: PermissionRegistry = ctx.request_context.lifespan_context["permissions"]
    return await compose_find(client=client, permissions=permissions, query=query, search_type=type)


@mcp.tool()
async def resolve(
    ctx: Context,
    url: Annotated[str, "An M365 URL: Teams message/meeting link, Outlook deep link, SharePoint page, or OneDrive file"],
) -> str:
    """Resolve any Microsoft 365 URL and return its content."""
    client: GraphClient = ctx.request_context.lifespan_context["client"]
    permissions: PermissionRegistry = ctx.request_context.lifespan_context["permissions"]
    return await compose_resolve(client=client, permissions=permissions, url=url)


def main():
    mcp.run()


if __name__ == "__main__":
    main()
