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
    "page": ["listItem"],
}

_DEFAULT_ENTITY_TYPES = ["message", "driveItem", "listItem"]


async def compose_find(
    client: GraphClient,
    permissions: PermissionRegistry,
    query: str,
    search_type: str | None,
) -> str:
    if search_type == "message":
        return await _search_chat_messages(client, query)

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


_MESSAGE_SEMAPHORE_LIMIT = 5
_MAX_CHATS_TO_SEARCH = 20
_MAX_HITS = 10


async def _search_chat_messages(client: GraphClient, query: str) -> str:
    """Search user's chat messages via enumeration.

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
        return format_section_error("Find", _error_reason(exc))
    if not chats:
        return format_search_results_markdown(query, [])

    chats, matched_words = _prefilter_chats_by_query(chats, query)
    chats = chats[:_MAX_CHATS_TO_SEARCH]

    all_words = [w for w in query.split() if len(w) >= 3]
    needles = [w for w in all_words if w.lower() not in matched_words]
    if not all_words:
        stripped = query.strip()
        needles = [stripped] if stripped else []

    semaphore = asyncio.Semaphore(_MESSAGE_SEMAPHORE_LIMIT)

    async def _bounded_fetch(chat_id: str) -> list[dict]:
        async with semaphore:
            if not needles:
                return await _recent_chat_messages(client, chat_id)
            return await _fetch_chat_messages(client, chat_id, needles)

    results = await asyncio.gather(
        *[_bounded_fetch(c["id"]) for c in chats],
        return_exceptions=True,
    )

    hits: list[dict] = []
    for result in results:
        if isinstance(result, list):
            hits.extend(result)

    hits.sort(key=lambda m: m.get("createdDateTime") or "", reverse=True)
    hits = hits[:_MAX_HITS]

    return format_search_results_markdown(query, [_to_search_hit(m) for m in hits])


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
        },
    }


async def _search_single(client: GraphClient, query: str, entity_types: list[str]) -> str:
    request: dict = {
        "entityTypes": entity_types,
        "query": {"queryString": query},
        "from": 0,
        "size": 10,
    }
    fields = _fields_for(entity_types)
    if fields:
        request["fields"] = fields
    payload = {"requests": [request]}

    try:
        response = await client.post("/search/query", payload)
    except GraphAPIError as exc:
        reason = _error_reason(exc)
        if "ChannelMessage" in exc.message or "ChatMessage" in exc.message:
            reason = (
                "Graph search for chat messages requires admin-consent "
                "ChannelMessage.Read.All, which this app doesn't have. "
                "Alternatives that work with the current Chat.Read scope: "
                "resolve(<chat URL>) for a specific chat's history; "
                "people(query='<name>') for recent chats/mail with someone; "
                "whats_new(scope='teams', since=...) for recent Teams activity"
            )
        return format_section_error("Find", reason)

    hits = _extract_hits(response)
    return format_search_results_markdown(query, hits)


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


async def _list_user_chats(client: GraphClient) -> list[dict]:
    """List user's chats, sorted by last-message recency (newest first).

    Returns up to 50 chats. Raises GraphAPIError on failure — caller
    formats an error message rather than confusing "no chats" with
    "call failed".
    """
    response = await client.get("/me/chats", params={
        "$expand": "members,lastMessagePreview",
        "$top": "50",
    })
    chats = (response or {}).get("value", [])
    chats.sort(
        key=lambda c: (c.get("lastMessagePreview") or {}).get("createdDateTime") or "",
        reverse=True,
    )
    return chats


def _prefilter_chats_by_query(chats: list[dict], query: str) -> tuple[list[dict], set[str]]:
    """Narrow chats to those whose members plausibly match the query.

    Returns (chats, matched_words). matched_words is the set of query words
    that positively matched at least one member displayName in the resulting
    chats. If no member matched any query word, the fallback returns
    (all_chats, set()).
    """
    words = [w.lower() for w in query.split() if len(w) >= 3]
    if not words:
        return chats, set()

    matched: list[dict] = []
    matched_words: set[str] = set()
    for chat in chats:
        member_names = " ".join(
            (m.get("displayName") or "").lower()
            for m in chat.get("members") or []
        )
        chat_matched = False
        for word in words:
            if word in member_names:
                chat_matched = True
                matched_words.add(word)
        if chat_matched:
            matched.append(chat)
    if matched:
        return matched, matched_words
    return chats, set()


async def _fetch_chat_messages(
    client: GraphClient,
    chat_id: str,
    needles: list[str],
    limit: int = 50,
) -> list[dict]:
    """Fetch chat messages and filter client-side against a set of needles.

    A message matches if ANY needle appears (case-insensitive substring) in
    its HTML-stripped, entity-decoded body. `_chat_id` is added to each hit.
    Empty needles or empty stripped set returns [] without fetching.
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
            hits.append({**msg, "_chat_id": chat_id})
    return hits


async def _recent_chat_messages(
    client: GraphClient,
    chat_id: str,
    limit: int = 20,
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
        hits.append({**msg, "_chat_id": chat_id})
    return hits
