"""find_v1 implementation — wraps composers.find.compose_find."""

from __future__ import annotations

from fastmcp import Context

from ...composers.find import compose_find
from .._helpers import _get_deps, wrap_errors
from .schemas import EmailHit, FileHit, FindPayload, FindResults, MessageHit, PageHit

TOOL_NAME = "find_v1"

_KIND_TO_MODEL = {
    "email": EmailHit,
    "file": FileHit,
    "message": MessageHit,
    "page": PageHit,
}


@wrap_errors(TOOL_NAME)
async def _find_v1_impl(ctx: Context, payload: FindPayload) -> FindResults:
    """Call the underlying composer and return a typed FindResults."""
    _, client, permissions = _get_deps(ctx)
    data, markdown = await compose_find(
        client=client,
        permissions=permissions,
        query=payload.query,
        search_type=payload.entity_type,
    )

    hits = []
    for h in data.get("hits", []):
        kind = h.get("kind")
        model_cls = _KIND_TO_MODEL.get(kind)
        if model_cls:
            try:
                hits.append(model_cls.model_validate(h))
            except Exception:
                pass

    return FindResults(
        query=payload.query,
        hits=hits,
        rendered_markdown=markdown,
    )
