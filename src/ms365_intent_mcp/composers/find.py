"""find composer — Microsoft Search API with per-type requests."""

import asyncio

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

    if len(entity_types) == 1:
        return await _search_single(client, query, entity_types)

    results = await asyncio.gather(
        *[_search_single_raw(client, query, [et]) for et in entity_types],
        return_exceptions=True,
    )

    hits = []
    for result in results:
        if isinstance(result, list):
            hits.extend(result)

    return format_search_results_markdown(query, hits)


async def _search_single(client: GraphClient, query: str, entity_types: list[str]) -> str:
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
        return format_section_error("Find", _error_reason(exc))

    hits = _extract_hits(response)
    return format_search_results_markdown(query, hits)


async def _search_single_raw(client: GraphClient, query: str, entity_types: list[str]) -> list[dict]:
    payload = {
        "requests": [
            {
                "entityTypes": entity_types,
                "query": {"queryString": query},
                "from": 0,
                "size": 5,
            }
        ]
    }
    response = await client.post("/search/query", payload)
    return _extract_hits(response)


def _extract_hits(response: dict) -> list[dict]:
    hits = []
    for result_set in (response or {}).get("value", []):
        for container in result_set.get("hitsContainers", []):
            for hit in container.get("hits", []):
                hits.append(hit)
    return hits
