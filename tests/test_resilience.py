"""Tests for CircuitBreaker state machine."""

import time
from unittest.mock import patch

import httpx
import pytest

from ms365_intent_mcp.graph import GraphAPIError
from ms365_intent_mcp.resilience import CircuitBreaker, CircuitOpenError, CircuitState


async def _raise_5xx():
    response = httpx.Response(500, request=httpx.Request("GET", "https://example.com"))
    raise httpx.HTTPStatusError("Server Error", request=response.request, response=response)


async def _raise_status(code: int):
    response = httpx.Response(code, request=httpx.Request("GET", "https://example.com"))
    raise httpx.HTTPStatusError(f"HTTP {code}", request=response.request, response=response)


async def _raise_graph_api_error(status: int):
    raise GraphAPIError(status, "Error", "test error")


async def _ok():
    return "ok"


class TestCircuitBreakerClosed:
    @pytest.mark.asyncio
    async def test_passes_through_on_success(self):
        cb = CircuitBreaker(failure_threshold=3)
        result = await cb.call(_ok)
        assert result == "ok"
        assert cb.state == CircuitState.CLOSED

    @pytest.mark.asyncio
    async def test_stays_closed_below_threshold(self):
        cb = CircuitBreaker(failure_threshold=3)
        for _ in range(2):
            with pytest.raises(httpx.HTTPStatusError):
                await cb.call(_raise_5xx)
        assert cb.state == CircuitState.CLOSED

    @pytest.mark.asyncio
    async def test_opens_at_threshold(self):
        cb = CircuitBreaker(failure_threshold=3)
        for _ in range(3):
            with pytest.raises(httpx.HTTPStatusError):
                await cb.call(_raise_5xx)
        assert cb.state == CircuitState.OPEN


class TestCircuitBreakerOpen:
    @pytest.mark.asyncio
    async def test_rejects_calls_when_open(self):
        cb = CircuitBreaker(failure_threshold=1, recovery_timeout=60.0)
        with pytest.raises(httpx.HTTPStatusError):
            await cb.call(_raise_5xx)
        with pytest.raises(CircuitOpenError):
            await cb.call(_ok)


class TestCircuitBreakerIgnoresNon5xx:
    @pytest.mark.asyncio
    async def test_404_does_not_trip(self):
        cb = CircuitBreaker(failure_threshold=1)
        with pytest.raises(httpx.HTTPStatusError):
            await cb.call(lambda: _raise_status(404))
        assert cb.state == CircuitState.CLOSED


class TestCircuitBreakerGraphAPIError:
    @pytest.mark.asyncio
    async def test_graph_api_500_trips_breaker(self):
        cb = CircuitBreaker(failure_threshold=2)
        for _ in range(2):
            with pytest.raises(GraphAPIError):
                await cb.call(lambda: _raise_graph_api_error(500))
        assert cb.state == CircuitState.OPEN

    @pytest.mark.asyncio
    async def test_graph_api_400_does_not_trip(self):
        cb = CircuitBreaker(failure_threshold=1)
        with pytest.raises(GraphAPIError):
            await cb.call(lambda: _raise_graph_api_error(400))
        assert cb.state == CircuitState.CLOSED

    @pytest.mark.asyncio
    async def test_graph_api_503_trips_breaker(self):
        cb = CircuitBreaker(failure_threshold=1)
        with pytest.raises(GraphAPIError):
            await cb.call(lambda: _raise_graph_api_error(503))
        assert cb.state == CircuitState.OPEN


class TestCircuitBreakerHalfOpen:
    @pytest.mark.asyncio
    async def test_transitions_to_half_open_after_timeout(self):
        cb = CircuitBreaker(failure_threshold=1, recovery_timeout=0.1)
        with pytest.raises(GraphAPIError):
            await cb.call(lambda: _raise_graph_api_error(500))
        assert cb.state == CircuitState.OPEN

        with patch("time.monotonic", return_value=time.monotonic() + 0.2):
            result = await cb.call(_ok)
        assert result == "ok"
        assert cb.state == CircuitState.CLOSED

    @pytest.mark.asyncio
    async def test_half_open_failure_reopens(self):
        cb = CircuitBreaker(failure_threshold=1, recovery_timeout=0.1)
        with pytest.raises(GraphAPIError):
            await cb.call(lambda: _raise_graph_api_error(500))
        assert cb.state == CircuitState.OPEN

        with patch("time.monotonic", return_value=time.monotonic() + 0.2):
            with pytest.raises(GraphAPIError):
                await cb.call(lambda: _raise_graph_api_error(500))
        assert cb.state == CircuitState.OPEN
