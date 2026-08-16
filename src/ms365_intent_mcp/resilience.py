"""Circuit breaker for Microsoft Graph API calls."""

from __future__ import annotations

import asyncio
import logging
import time
from collections.abc import Callable, Coroutine
from enum import Enum
from typing import TypeVar

import httpx

T = TypeVar("T")
_logger = logging.getLogger("ms365_intent_mcp")


class CircuitState(Enum):
    CLOSED = "closed"
    OPEN = "open"
    HALF_OPEN = "half_open"


class CircuitOpenError(Exception):
    def __init__(self, name: str, retry_after: float) -> None:
        self.name = name
        self.retry_after = retry_after
        remaining = max(0.0, retry_after - time.monotonic())
        super().__init__(f"Circuit '{name}' is OPEN. Retry in {remaining:.0f}s.")


def _is_server_error(exc: BaseException) -> bool:
    from .graph import GraphAPIError

    if isinstance(exc, GraphAPIError):
        return exc.status_code >= 500
    if isinstance(exc, httpx.HTTPStatusError):
        return exc.response.status_code >= 500
    return False


class CircuitBreaker:
    def __init__(
        self,
        name: str = "ms365_intent",
        failure_threshold: int = 5,
        recovery_timeout: float = 60.0,
        half_open_max_calls: int = 1,
        should_trip: Callable[[BaseException], bool] = _is_server_error,
    ) -> None:
        self.name = name
        self.failure_threshold = failure_threshold
        self.recovery_timeout = recovery_timeout
        self.half_open_max_calls = half_open_max_calls
        self.should_trip = should_trip
        self._state = CircuitState.CLOSED
        self._failure_count = 0
        self._opened_at: float = 0.0
        self._half_open_calls = 0
        self._lock = asyncio.Lock()

    @property
    def state(self) -> CircuitState:
        return self._state

    async def call(self, fn: Callable[[], Coroutine[None, None, T]]) -> T:
        async with self._lock:
            self._maybe_transition()
            if self._state == CircuitState.OPEN:
                raise CircuitOpenError(self.name, self._opened_at + self.recovery_timeout)
            if self._state == CircuitState.HALF_OPEN:
                if self._half_open_calls >= self.half_open_max_calls:
                    raise CircuitOpenError(self.name, self._opened_at + self.recovery_timeout)
                self._half_open_calls += 1

        try:
            result = await fn()
        except BaseException as exc:
            async with self._lock:
                if self._state == CircuitState.HALF_OPEN:
                    self._half_open_calls = max(0, self._half_open_calls - 1)
                if self.should_trip(exc):
                    self._on_failure()
            raise

        async with self._lock:
            if self._state == CircuitState.HALF_OPEN:
                self._half_open_calls = max(0, self._half_open_calls - 1)
            self._on_success()
        return result

    def _maybe_transition(self) -> None:
        if self._state == CircuitState.OPEN:
            if time.monotonic() >= self._opened_at + self.recovery_timeout:
                self._state = CircuitState.HALF_OPEN
                self._half_open_calls = 0

    def _on_failure(self) -> None:
        if self._state == CircuitState.HALF_OPEN:
            self._state = CircuitState.OPEN
            self._opened_at = time.monotonic()
            return
        self._failure_count += 1
        if self._failure_count >= self.failure_threshold:
            self._state = CircuitState.OPEN
            self._opened_at = time.monotonic()

    def _on_success(self) -> None:
        if self._state == CircuitState.HALF_OPEN:
            self._state = CircuitState.CLOSED
            self._failure_count = 0
            self._opened_at = 0.0
            return
        if self._failure_count > 0:
            self._failure_count = 0
