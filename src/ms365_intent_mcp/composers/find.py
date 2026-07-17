"""find composer — Microsoft Search API with per-type requests."""

import asyncio
import html

from ..formatters import _strip_teams_html, format_search_results_markdown, format_section_error
from ..graph import GraphClient, GraphAPIError
from ..permissions import PermissionRegistry
from ._utils import _error_reason, _list_user_chats, _prefilter_chats_by_query

_TYPE_MAP = {
    "email": ["message"],
    "file": ["driveItem"],
    "page": ["listItem"],
}

_DEFAULT_ENTITY_TYPES = ["message", "driveItem", "listItem"]


async def compose_find(
    client: GraphClient,
    permissions: PermissionRegistry,
    query: str,
    search_type: str | None,
) -> tuple[dict, str]:
    if search_type == "message":
        markdown, hits = await _search_chat_messages_with_hits(client, query)
        return {"query": query, "hits": hits}, markdown

    entity_types = _TYPE_MAP.get(search_type or "", _DEFAULT_ENTITY_TYPES)

    if len(entity_types) == 1:
        hits = await _search_single_raw(client, query, entity_types)
        markdown = format_search_results_markdown(query, hits)
    else:
        results = await asyncio.gather(
            *[_search_single_raw(client, query, [et]) for et in entity_types],
            return_exceptions=True,
        )
        hits = []
        for result in results:
            if isinstance(result, list):
                hits.extend(result)
        markdown = format_search_results_markdown(query, hits)

    structured_hits = [_hit_to_structured(h) for h in hits if _hit_to_structured(h) is not None]
    return {"query": query, "hits": structured_hits}, markdown


_MESSAGE_SEMAPHORE_LIMIT = 5
_MAX_CHATS_TO_SEARCH = 20
_MAX_HITS = 10


def _to_search_hit(msg: dict) -> dict:
    """Adapt a chatMessage dict into the hit shape expected by
    format_search_results_markdown.

    The formatter dispatches by resource['@odata.type'] substring. Preserve
    the message's own body/from so the existing "chatMessage" branch renders
    naturally.
    """
    return {
        "hitId": msg.get("id", ""),
        "rank": 0,
        "summary": "",
        "resource": {
            "@odata.type": "#microsoft.graph.chatMessage",
            "from": msg.get("from") or {},
            "body": msg.get("body") or {},
            "createdDateTime": msg.get("createdDateTime", ""),
            "chat_id": msg.get("_chat_id", ""),
            "chat_url": msg.get("_chat_web_url", ""),
        },
    }


def _hit_to_structured(hit: dict) -> dict | None:
    """Convert a Graph Search API hit dict into a structured dict for v1 impls.

    Returns None for unrecognised resource types (caller should filter these out).
    """
    resource = hit.get("resource") or {}
    odata_type = resource.get("@odata.type", "")
    if "chatMessage" in odata_type:
        sender = (resource.get("from") or {}).get("user", {}).get("displayName", "?")
        body_html = (resource.get("body") or {}).get("content", "")
        return {
            "kind": "message",
            "sender": sender,
            "body_preview": _strip_teams_html(body_html)[:200],
            "created": resource.get("createdDateTime"),
            "chat_id": resource.get("chat_id", ""),
            "chat_url": resource.get("chat_url", ""),
        }
    elif "message" in odata_type:
        return {
            "kind": "email",
            "subject": resource.get("subject", ""),
            "sender": (resource.get("from") or {}).get("emailAddress", {}).get("name", ""),
            "body_preview": resource.get("bodyPreview") or "",
            "web_link": resource.get("webLink"),
        }
    elif "driveItem" in odata_type:
        return {
            "kind": "file",
            "name": resource.get("name", ""),
            "web_url": resource.get("webUrl"),
            "size": resource.get("size"),
        }
    elif "listItem" in odata_type:
        fields = resource.get("fields") or {}
        return {
            "kind": "page",
            "title": fields.get("Title", resource.get("name", "")),
            "web_url": resource.get("webUrl"),
        }
    return None


async def _search_chat_messages_with_hits(client: GraphClient, query: str) -> tuple[str, list[dict]]:
    """Search user's chat messages via enumeration, returning markdown + structured hits.

    Delegated-scope alternative to POST /search/query (which requires
    admin-consent ChannelMessage.Read.All). Lists /me/chats, prefilters by
    member displayName when the query looks person-shaped, then fetches
    messages in parallel with a bounded semaphore.

    Content matching uses each significant word (3+ chars) as an independent
    needle so that "Diana second brain" matches messages containing "second brain"
    even when "diana" matches the chat member name rather than message content.
    Words that already matched a chat member's displayName are dropped from
    the needle set so a person-only query like "Diana" returns her recent
    messages instead of demanding her name appear in message bodies.
    """
    try:
        chats = await _list_user_chats(client)
    except GraphAPIError as exc:
        return format_section_error("Find", _error_reason(exc)), []
    if not chats:
        return format_search_results_markdown(query, []), []

    chats, matched_words = _prefilter_chats_by_query(chats, query)
    chats = chats[:_MAX_CHATS_TO_SEARCH]

    all_words = [w for w in query.split() if len(w) >= 3]
    needles = [w for w in all_words if w.lower() not in matched_words]
    if not all_words:
        stripped = query.strip()
        needles = [stripped] if stripped else []

    semaphore = asyncio.Semaphore(_MESSAGE_SEMAPHORE_LIMIT)

    async def _bounded_fetch(chat: dict) -> list[dict]:
        async with semaphore:
            chat_id = chat["id"]
            chat_web_url = chat.get("webUrl", "")
            if not needles:
                return await _recent_chat_messages(client, chat_id, chat_web_url=chat_web_url)
            return await _fetch_chat_messages(client, chat_id, needles, chat_web_url=chat_web_url)

    results = await asyncio.gather(
        *[_bounded_fetch(c) for c in chats],
        return_exceptions=True,
    )

    hits: list[dict] = []
    for result in results:
        if isinstance(result, list):
            hits.extend(result)

    hits.sort(key=lambda m: m.get("createdDateTime") or "", reverse=True)
    hits = hits[:_MAX_HITS]

    search_hits = [_to_search_hit(m) for m in hits]
    markdown = format_search_results_markdown(query, search_hits)
    structured = [_hit_to_structured(h) for h in search_hits]
    structured = [s for s in structured if s is not None]
    return markdown, structured


async def _search_single_raw(client: GraphClient, query: str, entity_types: list[str]) -> list[dict]:
    request: dict = {
        "entityTypes": entity_types,
        "query": {"queryString": query},
        "from": 0,
        "size": 5,
    }
    fields = _fields_for(entity_types)
    if fields:
        request["fields"] = fields
    payload = {"requests": [request]}
    response = await client.post("/search/query", payload)
    return _extract_hits(response)


def _fields_for(entity_types: list[str]) -> list[str]:
    """Return the fields projection for a Search request, if any is needed.

    Explicit fields keep webLink and bodyPreview stable for the "email" entity
    so the caller has both a snippet and a follow-up URL for resolve(). For
    driveItem/listItem, Graph's default projection is already sufficient and
    a mixed fields list would strip useful defaults on those types.
    """
    if entity_types == ["message"]:
        return [
            "id",
            "subject",
            "webLink",
            "bodyPreview",
            "from",
            "receivedDateTime",
        ]
    return []


def _extract_hits(response: dict) -> list[dict]:
    hits = []
    for result_set in (response or {}).get("value", []):
        for container in result_set.get("hitsContainers", []):
            for hit in container.get("hits", []):
                hits.append(hit)
    return hits


async def _fetch_chat_messages(
    client: GraphClient,
    chat_id: str,
    needles: list[str],
    limit: int = 50,
    chat_web_url: str = "",
) -> list[dict]:
    """Fetch chat messages and filter client-side against a set of needles.

    A message matches if ANY needle appears (case-insensitive substring) in
    its HTML-stripped, entity-decoded body. `_chat_id` and `_chat_web_url`
    are added to each hit. Empty needles or empty stripped set returns []
    without fetching.
    """
    lowered = [n.strip().lower() for n in needles if n and n.strip()]
    if not lowered:
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
        text = html.unescape(_strip_teams_html(body_html)).lower()
        if any(n in text for n in lowered):
            hits.append({**msg, "_chat_id": chat_id, "_chat_web_url": chat_web_url})
    return hits


async def _recent_chat_messages(
    client: GraphClient,
    chat_id: str,
    limit: int = 20,
    chat_web_url: str = "",
) -> list[dict]:
    """Return the N most recent messages from a chat, unfiltered.

    Used when the query resolves entirely to person-name words that already
    matched via the member prefilter — the intent then becomes "recent
    messages with this person", not substring search.
    """
    try:
        response = await client.get(f"/chats/{chat_id}/messages", params={
            "$top": str(limit),
        })
    except GraphAPIError:
        return []

    hits: list[dict] = []
    for msg in (response or {}).get("value", []):
        body_html = (msg.get("body") or {}).get("content", "")
        if not _strip_teams_html(body_html):
            continue
        hits.append({**msg, "_chat_id": chat_id, "_chat_web_url": chat_web_url})
    return hits
