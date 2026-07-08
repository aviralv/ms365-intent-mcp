"""people_v1 implementation — wraps composers.people.compose_people.

Structured fields stay placeholder-empty until Task 12 refactors
composers to return ``(dict, markdown)`` instead of markdown-only.
"""

from __future__ import annotations

from fastmcp import Context

from ...composers.people import compose_people
from .._helpers import _get_deps, wrap_errors
from .schemas import PeoplePayload, PersonDetail

TOOL_NAME = "people_v1"


@wrap_errors(TOOL_NAME)
async def _people_v1_impl(ctx: Context, payload: PeoplePayload) -> PersonDetail:
    """Look up a person and return a typed response.

    ``compose_people`` currently returns markdown only. Structured fields
    are set to best-guess / placeholder values; Task 12 will populate them
    from the refactored composer's dict output.
    """
    _, client, permissions = _get_deps(ctx)
    markdown = await compose_people(client, permissions, payload.query)
    return PersonDetail(
        name=payload.query,  # best guess until Task 12
        recent_mail=[],
        rendered_markdown=markdown,
    )
