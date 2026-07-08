"""find_v1 implementation — wraps composers.find.compose_find.

Structured hits are empty until Task 12 refactors composers to return
``(dict, markdown)`` tuples. ``rendered_markdown`` carries the current
composer output verbatim.
"""

from __future__ import annotations

from fastmcp import Context

from ...composers.find import compose_find
from .._helpers import _get_deps, wrap_errors
from .schemas import FindPayload, FindResults

TOOL_NAME = "find_v1"


@wrap_errors(TOOL_NAME)
async def _find_v1_impl(ctx: Context, payload: FindPayload) -> FindResults:
    """Call the underlying composer and return a typed FindResults.

    Structured hits are stubs — Task 12 fills them when the composer
    returns a ``(dict, markdown)`` tuple.
    """
    _, client, permissions = _get_deps(ctx)
    markdown = await compose_find(
        client=client,
        permissions=permissions,
        query=payload.query,
        search_type=payload.entity_type,
    )
    return FindResults(
        query=payload.query,
        hits=[],
        rendered_markdown=markdown,
    )
