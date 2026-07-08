"""whats_new implementation — wraps composers.whats_new.compose_whats_new."""

from __future__ import annotations

from fastmcp import Context

from ...composers.whats_new import compose_whats_new
from .._helpers import _get_deps, wrap_errors
from .schemas import EventSummary, MailItem, TeamsItem, WhatsNewPayload, WhatsNewSummary

TOOL_NAME = "whats_new"


@wrap_errors(TOOL_NAME)
async def _whats_new_impl(ctx: Context, payload: WhatsNewPayload) -> WhatsNewSummary:
    """Call the underlying composer and return a typed WhatsNewSummary."""
    config, client, permissions = _get_deps(ctx)
    composer_scope = None if payload.scope == "all" else payload.scope
    data, markdown = await compose_whats_new(
        client=client,
        permissions=permissions,
        since=payload.since.isoformat(),
        scope=composer_scope,
        timezone=config.default_timezone,
    )

    mail = []
    for m in data.get("mail", []):
        try:
            mail.append(MailItem.model_validate(m))
        except Exception:
            pass

    events = []
    for e in data.get("events", []):
        try:
            events.append(EventSummary.model_validate(e))
        except Exception:
            pass

    teams = []
    for t in data.get("teams", []):
        try:
            teams.append(TeamsItem.model_validate(t))
        except Exception:
            pass

    return WhatsNewSummary(
        since=payload.since,
        mail=mail,
        events=events,
        teams=teams,
        rendered_markdown=markdown,
    )
