"""whats_new_v1 implementation — wraps composers.whats_new.compose_whats_new.

Structured fields (mail, events, teams) are empty until Task 12 refactors
composers to return ``(dict, markdown)`` tuples. ``rendered_markdown``
carries the current composer output verbatim.
"""

from __future__ import annotations

from fastmcp import Context

from ...composers.whats_new import compose_whats_new
from .._helpers import _get_deps, wrap_errors
from .schemas import WhatsNewPayload, WhatsNewSummary

TOOL_NAME = "whats_new_v1"


@wrap_errors(TOOL_NAME)
async def _whats_new_v1_impl(ctx: Context, payload: WhatsNewPayload) -> WhatsNewSummary:
    """Call the underlying composer and return a typed WhatsNewSummary.

    Structured fields (mail, events, teams) are stubs — Task 12 fills them
    when the composer returns a ``(dict, markdown)`` tuple.

    scope='all' is translated to None for the underlying composer (which
    treats None as 'all' internally).
    """
    config, client, permissions = _get_deps(ctx)
    composer_scope = None if payload.scope == "all" else payload.scope
    markdown = await compose_whats_new(
        client=client,
        permissions=permissions,
        since=payload.since.isoformat(),
        scope=composer_scope,
        timezone=config.default_timezone,
    )
    return WhatsNewSummary(
        since=payload.since,
        mail=[],  # Task 12 fills these when composer returns (dict, markdown)
        events=[],
        teams=[],
        rendered_markdown=markdown,
    )
