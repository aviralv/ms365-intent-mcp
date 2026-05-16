"""Shared utilities for composers."""

from ..graph import GraphAPIError

NOISE_PATTERNS = [
    "noreply@", "no-reply@", "notifications@", "mailer@",
    "newsletter@", "digest@", "productboard", "stackoverflow",
    "github.com", "jira", "confluence", "atlassian", "slack",
    "successfactors", "concur", "workday",
]


def _error_reason(exc: Exception) -> str:
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
