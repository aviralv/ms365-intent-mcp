"""people_v1 implementation — wraps composers.people.compose_people."""

from __future__ import annotations

from fastmcp import Context

from ...composers.people import compose_people
from .._helpers import _get_deps, wrap_errors
from .schemas import MailPreview, ChatPreview, PeoplePayload, PersonDetail

TOOL_NAME = "people_v1"


@wrap_errors(TOOL_NAME)
async def _people_v1_impl(ctx: Context, payload: PeoplePayload) -> PersonDetail:
    """Look up a person and return a typed PersonDetail response."""
    _, client, permissions = _get_deps(ctx)
    data, markdown = await compose_people(client, permissions, payload.query)
    return PersonDetail(
        name=data["name"],
        email=data["email"] or None,
        job_title=data.get("job_title"),
        recent_mail=[MailPreview.model_validate(m) for m in data["recent_mail"]],
        recent_chat=ChatPreview.model_validate(data["recent_chat"]) if data.get("recent_chat") else None,
        rendered_markdown=markdown,
    )
