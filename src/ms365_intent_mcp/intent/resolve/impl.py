"""resolve_v1 implementation — dispatches URL to compose_resolve and wraps
the markdown-only return into a typed ResolvedContent response.

Structured fields on content types are placeholder stubs until Task 12
refactors composers to return ``(dict, markdown)`` tuples.
"""

from __future__ import annotations

from fastmcp import Context

from ...composers.resolve import compose_resolve
from .._helpers import IntentError, _get_deps, wrap_errors
from .schemas import (
    ChannelMessageContent,
    ChatMessageContent,
    ChatThreadContent,
    EmailContent,
    MeetingContent,
    OneDriveFileContent,
    ResolvedContent,
    ResolvePayload,
    SharePointPageContent,
)

TOOL_NAME = "resolve_v1"


@wrap_errors(TOOL_NAME)
async def _resolve_v1_impl(ctx: Context, payload: ResolvePayload) -> ResolvedContent:
    """Resolve an M365 URL and return a typed response.

    URL type is auto-detected by resolver.py's regex dispatch table.
    ``wrap_errors`` catches ``IntentError`` / ``GraphAPIError`` and returns
    an ``ErrorResponse``.
    """
    _, client, permissions = _get_deps(ctx)

    from ...resolver import UrlParseError, resolve_url

    try:
        resolved = resolve_url(str(payload.url))
    except UrlParseError as exc:
        raise IntentError("invalid_id", str(exc))

    url_type = resolved.url_type
    markdown = await compose_resolve(client, permissions, str(payload.url))

    if url_type == "email":
        data = EmailContent(kind="email", subject="(pending)", sender="(pending)", body="")
    elif url_type == "chat_thread":
        data = ChatThreadContent(kind="chat_thread")
    elif url_type == "chat_message":
        data = ChatMessageContent(kind="chat_message", sender="(pending)", body="")
    elif url_type == "channel_message":
        data = ChannelMessageContent(kind="channel_message", sender="(pending)", body="")
    elif url_type == "meeting":
        data = MeetingContent(kind="meeting", subject="(pending)")
    elif url_type == "sharepoint_page":
        data = SharePointPageContent(kind="sharepoint_page", title="(pending)")
    elif url_type in ("onedrive_file", "onedrive_share_link"):
        data = OneDriveFileContent(kind="onedrive_file", name="(pending)")
    else:
        raise IntentError("invalid_id", f"Unknown url_type: {url_type}")

    return ResolvedContent(
        url=payload.url,
        kind=data.kind,
        data=data,
        rendered_markdown=markdown,
    )
