"""people implementation — wraps composers.people.compose_people."""

from __future__ import annotations

from fastmcp import Context

from ...composers.people import compose_people
from .._helpers import _get_deps, wrap_errors
from .schemas import ChatPreview, MailPreview, OutOfOfficeStatus, PeoplePayload, PersonDetail

TOOL_NAME = "people"


@wrap_errors(TOOL_NAME)
async def _people_impl(ctx: Context, payload: PeoplePayload) -> PersonDetail:
    """Look up a person and return a typed PersonDetail response."""
    _, client, permissions = _get_deps(ctx)
    data, markdown = await compose_people(client, permissions, payload.query)

    ooo_raw = data.get("out_of_office")
    ooo = OutOfOfficeStatus.model_validate(ooo_raw) if ooo_raw else None

    return PersonDetail(
        name=data.get("name", ""),
        email=data.get("email") or None,
        job_title=data.get("job_title"),
        out_of_office=ooo,
        recent_mail=[
            MailPreview.model_validate(m) for m in data.get("recent_mail", [])
        ],
        recent_chat=(
            ChatPreview.model_validate(data["recent_chat"])
            if data.get("recent_chat")
            else None
        ),
        rendered_markdown=markdown,
    )
