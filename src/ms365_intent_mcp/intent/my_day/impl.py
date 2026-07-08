"""my_day_v1 implementation — wraps composers.my_day.compose_my_day.

Structured fields (events, mail, teams) are empty until Task 12 refactors
composers to return ``(dict, markdown)`` tuples. ``rendered_markdown``
carries the current composer output verbatim.
"""

from __future__ import annotations

from datetime import date

from fastmcp import Context

from ...composers.my_day import compose_my_day
from .._helpers import _get_deps, wrap_errors
from .schemas import MailSummary, MyDayPayload, MyDaySummary, TeamsActivitySummary

TOOL_NAME = "my_day_v1"


@wrap_errors(TOOL_NAME)
async def _my_day_v1_impl(ctx: Context, payload: MyDayPayload) -> MyDaySummary:
    """Call the underlying composer and return a typed MyDaySummary.

    Structured fields (events, mail, teams) are stubs — Task 12 fills them
    when the composer returns a ``(dict, markdown)`` tuple.
    """
    config, client, permissions = _get_deps(ctx)
    target_date = payload.date or date.today()
    markdown = await compose_my_day(
        client,
        permissions,
        target_date.strftime("%Y-%m-%d"),
        config.default_timezone,
    )
    return MyDaySummary(
        date=target_date,
        events=[],  # Task 12 fills these when composer returns (dict, markdown)
        mail=MailSummary(),
        teams=TeamsActivitySummary(),
        rendered_markdown=markdown,
    )
