"""compose implementation — dispatches on ComposePayload's ``type``.

Wraps the existing ``composers.compose.compose_action`` and adapts its
(dict, markdown) return into typed response models.
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
    EventForwarded,
    TeamsMessageSent,
)

TOOL_NAME = "compose"


@wrap_errors(TOOL_NAME)
async def _compose_impl(ctx: Context, payload: ComposePayload) -> ComposeResponse:
    """Dispatch a compose payload to the underlying composer.

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
            "to": [{"email": r.email, "name": r.name or r.email} for r in (payload.to or [])],
            "importance": payload.importance,
        }
        if payload.cc:
            params["cc"] = [{"email": r.email, "name": r.name or r.email} for r in payload.cc]
        data, markdown = await compose_action(
            client,
            permissions,
            ComposeType.EMAIL_DRAFT,
            params,
        )
    elif payload.mode == "forward":
        params = {
            "message_id": payload.in_reply_to_message_id,
            "body": payload.body,
            "to": [{"email": r.email, "name": r.name or r.email} for r in (payload.to or [])],
        }
        if payload.cc:
            params["cc"] = [{"email": r.email, "name": r.name or r.email} for r in payload.cc]
        data, markdown = await compose_action(
            client,
            permissions,
            ComposeType.EMAIL_FORWARD,
            params,
        )
    else:
        params = {
            "message_id": payload.in_reply_to_message_id,
            "body": payload.body,
            "reply_all": payload.mode == "reply_all",
        }
        if payload.comment:
            params["comment"] = payload.comment
        data, markdown = await compose_action(
            client,
            permissions,
            ComposeType.REPLY_DRAFT,
            params,
        )

    # Build recipient list from data (preferred) or payload (fallback for error paths)
    to_list = []
    for r in data.get("to", []):
        email = r.get("email", "")
        if email:
            to_list.append(Recipient(email=email, name=r.get("name") or None))
    if not to_list:
        to_list = payload.to or []

    return EmailDraftCreated(
        draft_id=data.get("draft_id") or "unknown",
        subject=data.get("subject") or payload.subject or "(reply)",
        to=to_list,
        web_link=data.get("web_link") or "https://outlook.office.com/mail/inbox",  # type: ignore[arg-type]
        rendered_markdown=markdown,
    )


async def _handle_event(
    payload: ComposeEvent,
    client,
    permissions,
    config_default_tz: str,
) -> EventCreated | EventForwarded:
    """Create a calendar event via the legacy composer, or forward an existing one."""
    if payload.mode == "forward":
        params_fwd: dict = {
            "event_id": payload.event_id,
            "to": [{"email": r.email, "name": r.name or r.email} for r in (payload.to or [])],
        }
        if payload.comment:
            params_fwd["comment"] = payload.comment
        data, markdown = await compose_action(
            client,
            permissions,
            ComposeType.EVENT_FORWARD,
            params_fwd,
        )
        to_list = [
            Recipient(email=r["email"], name=r.get("name") or None) for r in data.get("to", [])
        ]
        return EventForwarded(to=to_list, rendered_markdown=markdown)

    params: dict = {
        "subject": payload.subject,
        "start": payload.start.isoformat(),
        "end": payload.end.isoformat(),
        "timezone": payload.timezone,
        "is_online_meeting": payload.is_online_meeting,
    }
    if payload.body:
        params["body"] = payload.body
    if payload.location:
        params["location"] = payload.location
    if payload.attendees:
        params["attendees"] = [
            {"email": a.email, "name": a.name or a.email, "type": a.type} for a in payload.attendees
        ]

    data, markdown = await compose_action(client, permissions, ComposeType.EVENT, params)

    return EventCreated(
        event_id=data.get("event_id") or "unknown",
        subject=data.get("subject") or payload.subject,
        start=payload.start,
        end=payload.end,
        join_url=data.get("join_url") or None,
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
    data, markdown = await compose_action(client, permissions, ComposeType.TEAMS_MESSAGE, params)

    return TeamsMessageSent(
        message_id=data.get("message_id") or "unknown",
        chat_id=payload.chat_id,
        rendered_markdown=markdown,
    )
