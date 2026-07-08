"""Cross-tool helpers for the intent surface."""

from __future__ import annotations

import time
from functools import wraps
from typing import Any

from fastmcp import Context

from ..config import Config
from ..graph import GraphAPIError, GraphClient
from ..permissions import PermissionRegistry
from ._shared import ErrorResponse


class IntentError(Exception):
    """Raised by impls when a known error should surface as ``ErrorResponse``.

    ``code`` must match the ``Literal`` on ``ErrorResponse.code``. Unknown codes
    won't validate; keep the two lists aligned.
    """

    def __init__(self, code: str, message: str, retryable: bool = False) -> None:
        self.code = code
        self.message = message
        self.retryable = retryable
        super().__init__(message)


def _get_deps(ctx: Context) -> tuple[Config, GraphClient, PermissionRegistry]:
    """Extract the three dependencies from the FastMCP lifespan context."""
    lifespan = ctx.request_context.lifespan_context
    return lifespan["config"], lifespan["client"], lifespan["permissions"]


def wrap_errors(func_name: str):
    """Decorator that catches ``IntentError`` / ``GraphAPIError`` and returns ``ErrorResponse``.

    Composers raise ``IntentError`` for domain-level errors; ``GraphAPIError`` from
    the Graph client is treated as ``graph_api_error`` with retryable=True on 429/503.
    """

    def _decorator(func):
        @wraps(func)
        async def _wrapped(*args, **kwargs):
            try:
                return await func(*args, **kwargs)
            except IntentError as exc:
                return ErrorResponse(
                    code=exc.code,  # type: ignore[arg-type]
                    message=exc.message,
                    retryable=exc.retryable,
                )
            except GraphAPIError as exc:
                if exc.status_code in (429, 503):
                    return ErrorResponse(
                        code="rate_limited",
                        message=f"{exc.error_code}: {exc.message}",
                        retryable=True,
                    )
                return ErrorResponse(
                    code="graph_api_error",
                    message=f"{exc.error_code}: {exc.message}",
                    retryable=exc.status_code in (429, 503),
                )
            except Exception as exc:
                return ErrorResponse(
                    code="graph_api_error",
                    message=f"unhandled exception: {type(exc).__name__}: {exc}",
                    retryable=False,
                )

        _wrapped.__name__ = func_name
        return _wrapped

    return _decorator


# ---------------------------------------------------------------------------
# Idempotency cache — in-memory, per-process, 10-min TTL, 1000-entry cap.
# ---------------------------------------------------------------------------

_IDEMPOTENCY_TTL_SECONDS = 600
_IDEMPOTENCY_MAX_ENTRIES = 1000
_idempotency_cache: dict[str, tuple[float, Any]] = {}


def idempotency_lookup(tool_name: str, key: str | None) -> Any | None:
    """Return the cached response for ``(tool_name, key)``, or ``None`` if absent/expired.

    Silent no-op when ``key`` is falsy — idempotency is opt-in per call.
    """
    if not key:
        return None
    cache_key = f"{tool_name}:{key}"
    entry = _idempotency_cache.get(cache_key)
    if entry is None:
        return None
    stored_at, response = entry
    if time.time() - stored_at > _IDEMPOTENCY_TTL_SECONDS:
        _idempotency_cache.pop(cache_key, None)
        return None
    return response


def idempotency_store(tool_name: str, key: str | None, response: Any) -> None:
    """Store a response under ``(tool_name, key)``. No-op if ``key`` is falsy.

    Evicts the oldest entry when the cache is full — simple LRU-ish behavior
    keyed on ``stored_at`` timestamps, not a proper LRU. Good enough for a
    retry-dedup guard on a low-write MCP.
    """
    if not key:
        return
    cache_key = f"{tool_name}:{key}"
    if len(_idempotency_cache) >= _IDEMPOTENCY_MAX_ENTRIES:
        oldest = min(_idempotency_cache.items(), key=lambda kv: kv[1][0])
        _idempotency_cache.pop(oldest[0], None)
    _idempotency_cache[cache_key] = (time.time(), response)


def idempotency_clear() -> None:
    """Test helper — flush all cache entries."""
    _idempotency_cache.clear()
