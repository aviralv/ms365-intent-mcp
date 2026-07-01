"""find composer — Microsoft Search API with per-type requests."""

import asyncio
import html

from ..formatters import _strip_teams_html, format_search_results_markdown, format_section_error
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


async def _list_user_chats(client: GraphClient) -> list[dict]:
    """List user's chats, sorted by last-message recency (newest first).

    Returns up to 50 chats. Empty list on any GraphAPIError so callers can
    treat "no chats accessible" and "call failed" uniformly.
    """
    try:
        response = await client.get("/me/chats", params={
            "$expand": "members,lastMessagePreview",
            "$top": "50",
        })
    except GraphAPIError:
        return []
    chats = (response or {}).get("value", [])
    chats.sort(
        key=lambda c: (c.get("lastMessagePreview") or {}).get("createdDateTime") or "",
        reverse=True,
    )
    return chats


def _prefilter_chats_by_query(chats: list[dict], query: str) -> list[dict]:
    """Narrow chats to those whose members plausibly match the query.

    Heuristic: for each 3+ char word in the query, if any chat has a member
    whose displayName contains that word (case-insensitive), keep only the
    matching chats. If no chat matches any query word, return the full list
    (query wasn't person-shaped, or the person isn't in the top-N chats).
    """
    words = [w.lower() for w in query.split() if len(w) >= 3]
    if not words:
        return chats

    matched: list[dict] = []
    for chat in chats:
        member_names = " ".join(
            (m.get("displayName") or "").lower()
            for m in chat.get("members") or []
        )
        if any(word in member_names for word in words):
            matched.append(chat)
    return matched if matched else chats


async def _fetch_chat_messages(
    client: GraphClient,
    chat_id: str,
    query: str,
    limit: int = 50,
) -> list[dict]:
    """Fetch chat messages and filter client-side by query substring.

    Graph's /chats/{id}/messages supports neither $filter nor $search on body,
    so filtering is client-side. Case-insensitive substring match against the
    HTML-stripped body text. Empty list on GraphAPIError so callers can gather
    across chats without per-chat error handling.
    """
    needle = query.strip().lower()
    if not needle:
        return []
    try:
        response = await client.get(f"/chats/{chat_id}/messages", params={
            "$top": str(limit),
        })
    except GraphAPIError:
        return []

    hits: list[dict] = []
    for msg in (response or {}).get("value", []):
        body_html = (msg.get("body") or {}).get("content", "")
        text = _strip_teams_html(body_html)
        if needle in html.unescape(text).lower():
            hits.append({**msg, "_chat_id": chat_id})
    return hits
