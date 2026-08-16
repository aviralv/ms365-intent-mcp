"""Shared utilities for composers."""

from ..graph import GraphAPIError, GraphClient


def _escape_odata(value: str) -> str:
    """Escape a string for use in OData $filter expressions (doubles single quotes)."""
    return value.replace("'", "''")


NOISE_PATTERNS = [
    "noreply@",
    "no-reply@",
    "notifications@",
    "mailer@",
    "newsletter@",
    "digest@",
    "productboard",
    "stackoverflow",
    "github.com",
    "jira",
    "confluence",
    "atlassian",
    "slack",
    "successfactors",
    "concur",
    "workday",
]


def _error_reason(exc: BaseException) -> str:
    if isinstance(exc, GraphAPIError):
        if exc.status_code == 429:
            return "rate limited — retry shortly"
        if exc.status_code >= 500:
            return "Microsoft service error"
        return exc.message
    if "timeout" in str(exc).lower() or "timed out" in str(exc).lower():
        return "timed out"
    return str(exc)[:100]


def _is_noise(msg: dict) -> bool:
    from_addr = ((msg.get("from") or {}).get("emailAddress") or {}).get("address", "").lower()
    return any(p in from_addr for p in NOISE_PATTERNS)


def _sender_name(msg: dict) -> str:
    return ((msg.get("from") or {}).get("emailAddress") or {}).get("name", "Unknown")


def _chat_sender(preview: dict) -> str:
    from_field = preview.get("from") or {}
    user_field = from_field.get("user") or {}
    return user_field.get("displayName", "Unknown")


def _build_mail_summary(all_msgs: list[dict]) -> dict:
    """Build a structured mail summary from a list of messages.

    Returns dict with keys: all_count, relevant_count, high_importance, needs_attention.
    """
    relevant = [m for m in all_msgs if not _is_noise(m)]
    high_importance = [
        {"subject": m.get("subject", "?"), "from": _sender_name(m)}
        for m in relevant
        if m.get("importance") == "high"
    ]
    needs_attention = [
        {"subject": m.get("subject", "?"), "from": _sender_name(m)} for m in relevant[:5]
    ]
    return {
        "all_count": len(all_msgs),
        "relevant_count": len(relevant),
        "high_importance": high_importance[:5],
        "needs_attention": needs_attention,
    }


async def _list_user_chats(client: GraphClient) -> list[dict]:
    """List user's chats, sorted by last-message recency (newest first).

    Returns up to 50 chats. Raises GraphAPIError on failure — caller
    formats an error message rather than confusing "no chats" with
    "call failed".
    """
    response = await client.get(
        "/me/chats",
        params={
            "$expand": "members,lastMessagePreview",
            "$top": "50",
        },
    )
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
            (m.get("displayName") or "").lower() for m in chat.get("members") or []
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
