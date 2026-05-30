"""Shared utilities for composers."""

from ..graph import GraphAPIError


def _escape_odata(value: str) -> str:
    """Escape a string for use in OData $filter expressions (doubles single quotes)."""
    return value.replace("'", "''")

NOISE_PATTERNS = [
    "noreply@", "no-reply@", "notifications@", "mailer@",
    "newsletter@", "digest@", "productboard", "stackoverflow",
    "github.com", "jira", "confluence", "atlassian", "slack",
    "successfactors", "concur", "workday",
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
        for m in relevant if m.get("importance") == "high"
    ]
    needs_attention = [
        {"subject": m.get("subject", "?"), "from": _sender_name(m)}
        for m in relevant[:5]
    ]
    return {
        "all_count": len(all_msgs),
        "relevant_count": len(relevant),
        "high_importance": high_importance[:5],
        "needs_attention": needs_attention,
    }
