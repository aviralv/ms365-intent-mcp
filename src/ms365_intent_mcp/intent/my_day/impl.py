"""my_day_v1 implementation — wraps composers.my_day.compose_my_day."""

from __future__ import annotations

from datetime import date

from fastmcp import Context

from ...composers.my_day import compose_my_day
from .._helpers import _get_deps, wrap_errors
from .schemas import EventSummary, MailSummary, MyDayPayload, MyDaySummary, TeamsActivitySummary

TOOL_NAME = "my_day_v1"


@wrap_errors(TOOL_NAME)
async def _my_day_v1_impl(ctx: Context, payload: MyDayPayload) -> MyDaySummary:
    """Call the underlying composer and return a typed MyDaySummary."""
    config, client, permissions = _get_deps(ctx)
    target_date = payload.date or date.today()
    data, markdown = await compose_my_day(
        client,
        permissions,
        target_date.strftime("%Y-%m-%d"),
        config.default_timezone,
    )

    events = []
    for e in data.get("events", []):
        try:
            events.append(EventSummary.model_validate(e))
        except Exception:
            pass

    mail_data = data.get("mail", {})
    teams_data = data.get("teams", {})

    return MyDaySummary(
        date=target_date,
        events=events,
        mail=MailSummary(
            unread_count=mail_data.get("unread_count", 0),
            relevant_count=mail_data.get("relevant_count", 0),
            flagged_count=mail_data.get("flagged_count", 0),
        ),
        teams=TeamsActivitySummary(
            recent_message_count=teams_data.get("recent_message_count", 0),
        ),
        rendered_markdown=markdown,
    )
