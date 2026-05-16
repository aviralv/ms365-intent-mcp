"""find composer — Microsoft Search API with 403 chatMessage fallback."""

from ..formatters import format_search_results_markdown, format_section_error
from ..graph import GraphClient, GraphAPIError
from ..permissions import PermissionRegistry
from ._utils import _error_reason

_TYPE_MAP = {
    "email": ["message"],
    "file": ["driveItem"],
    "message": ["chatMessage"],
    "page": ["listItem"],
}

_DEFAULT_ENTITY_TYPES = ["message", "driveItem", "listItem"]


async def compose_find(
    client: GraphClient,
    permissions: PermissionRegistry,
    query: str,
    search_type: str | None,
) -> str:
    entity_types = _TYPE_MAP.get(search_type or "", _DEFAULT_ENTITY_TYPES)

    payload = {
        "requests": [
            {
                "entityTypes": entity_types,
                "query": {"queryString": query},
                "from": 0,
                "size": 10,
            }
        ]
    }

    try:
        response = await client.post("/search/query", payload)
    except GraphAPIError as exc:
        if exc.status_code == 403 and "chatMessage" in entity_types:
            entity_types_fallback = [t for t in entity_types if t != "chatMessage"]
            if not entity_types_fallback:
                return "### Find\nSearch unavailable — chatMessage scope not granted."
            payload["requests"][0]["entityTypes"] = entity_types_fallback
            try:
                response = await client.post("/search/query", payload)
            except GraphAPIError as exc2:
                return format_section_error("Find", _error_reason(exc2))
        else:
            return format_section_error("Find", _error_reason(exc))

    hits = _extract_hits(response)
    return format_search_results_markdown(query, hits)


def _extract_hits(response: dict) -> list[dict]:
    hits = []
    for result_set in (response or {}).get("value", []):
        for container in result_set.get("hitsContainers", []):
            for hit in container.get("hits", []):
                hits.append(hit)
    return hits
