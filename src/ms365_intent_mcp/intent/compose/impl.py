"""compose_v1 implementation — dispatches on ComposePayload's ``type``.

Wraps the existing ``composers.compose.compose_action`` and adapts its
markdown-only return into typed response models. Structured fields
(``draft_id``, ``event_id``, ``web_link``, ``join_url``) are placeholder
values until Task 12 refactors composers to return ``(dict, markdown)``.
"""

from __future__ import annotations

from fastmcp import Context

from ...composers.compose import ComposeType, compose_action
from .._helpers import (
    IntentError,
    _get_deps,
    idempotency_lookup,
    idempotency_store,
    wrap_errors,
)
from .._shared import Recipient
from .schemas import (
    ComposeEmail,
    ComposeEvent,
    ComposePayload,
    ComposeResponse,
    ComposeTeamsMessage,
    EmailDraftCreated,
    EventCreated,
    TeamsMessageSent,
)

TOOL_NAME = "compose_v1"

# Placeholder ID used until Task 12 makes composers return (dict, markdown).
# When a live-verify script sees this string in a response, it means the
# composer refactor hasn't landed for that variant yet.
_PENDING = "pending-composer-dict-refactor"
_PENDING_URL = "https://outlook.office.com/mail/inbox"


@wrap_errors(TOOL_NAME)
async def _compose_v1_impl(ctx: Context, payload: ComposePayload) -> ComposeResponse:
    """Dispatch a compose_v1 payload to the underlying composer.

    Returns a typed response model. On error, ``wrap_errors`` catches
    ``IntentError`` / ``GraphAPIError`` and returns an ``ErrorResponse``.
    Idempotency: if ``payload.idempotency_key`` is set and the same key was
    seen in the past 10 minutes, the cached response is returned without
    re-executing.
    """
    config, client, permissions = _get_deps(ctx)

    idem_key = getattr(payload, "idempotency_key", None)
    cached = idempotency_lookup(TOOL_NAME, idem_key)
    if cached is not None:
        return cached

    if isinstance(payload, ComposeEmail):
        response: ComposeResponse = await _handle_email(payload, client, permissions)
    elif isinstance(payload, ComposeEvent):
        response = await _handle_event(payload, client, permissions, config.default_timezone)
    elif isinstance(payload, ComposeTeamsMessage):
        response = await _handle_teams_message(payload, client, permissions)
    else:
        raise IntentError(
            "validation_error",
            f"Unknown compose variant: {type(payload).__name__}",
        )

    idempotency_store(TOOL_NAME, idem_key, response)
    return response


async def _handle_email(
    payload: ComposeEmail,
    client,
    permissions,
) -> EmailDraftCreated:
    """Route to email_draft or reply_draft based on ``mode``."""
    if payload.mode == "new":
        params: dict = {
            "subject": payload.subject,
            "body": payload.body,
            "to": [
                {"email": r.email, "name": r.name or r.email}
                for r in (payload.to or [])
            ],
            "importance": payload.importance,
        }
        if payload.cc:
            params["cc"] = [
                {"email": r.email, "name": r.name or r.email} for r in payload.cc
            ]
        markdown = await compose_action(
            client, permissions, ComposeType.EMAIL_DRAFT, params,
        )
    else:
        # reply / reply_all / forward — all route through the reply_draft path.
        # `forward` uses the same Graph endpoint (createReply/createReplyAll)
        # with reply_all=False, which is a light abuse of the Graph API but
        # matches the current legacy composer's behavior. Task 12 will refine.
        params = {
            "message_id": payload.in_reply_to_message_id,
            "body": payload.body,
            "reply_all": payload.mode == "reply_all",
        }
        if payload.comment:
            params["comment"] = payload.comment
        markdown = await compose_action(
            client, permissions, ComposeType.REPLY_DRAFT, params,
        )

    return EmailDraftCreated(
        draft_id=_PENDING,
        subject=payload.subject or "(reply)",
        to=payload.to or [Recipient(email="pending@example.com")],
        web_link=_PENDING_URL,  # type: ignore[arg-type]
        rendered_markdown=markdown,
    )


async def _handle_event(
    payload: ComposeEvent,
    client,
    permissions,
    config_default_tz: str,
) -> EventCreated:
    """Create a calendar event via the legacy composer."""
    params: dict = {
        "subject": payload.subject,
        "start": payload.start.isoformat(),
        "end": payload.end.isoformat(),
        "timezone": payload.timezone or config_default_tz,
        "is_online_meeting": payload.is_online_meeting,
    }
    if payload.body:
        params["body"] = payload.body
    if payload.location:
        params["location"] = payload.location
    if payload.attendees:
        params["attendees"] = [
            {"email": a.email, "name": a.name or a.email, "type": a.type}
            for a in payload.attendees
        ]

    markdown = await compose_action(client, permissions, ComposeType.EVENT, params)

    return EventCreated(
        event_id=_PENDING,
        subject=payload.subject,
        start=payload.start,
        end=payload.end,
        join_url=None,
        rendered_markdown=markdown,
    )


async def _handle_teams_message(
    payload: ComposeTeamsMessage,
    client,
    permissions,
) -> TeamsMessageSent:
    """Send a Teams chat message via the legacy composer."""
    params = {
        "chat_id": payload.chat_id,
        "content": payload.content,
        "content_type": payload.content_type,
    }
    markdown = await compose_action(client, permissions, ComposeType.TEAMS_MESSAGE, params)

    return TeamsMessageSent(
        message_id=_PENDING,
        chat_id=payload.chat_id,
        rendered_markdown=markdown,
    )
