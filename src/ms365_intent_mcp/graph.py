"""Async HTTP client for Microsoft Graph API."""

import asyncio
import logging
from collections.abc import Callable
from typing import Any

import httpx

from .resilience import CircuitBreaker

_logger = logging.getLogger("ms365_intent_mcp")


class GraphAPIError(Exception):
    def __init__(self, status_code: int, error_code: str, message: str):
        self.status_code = status_code
        self.error_code = error_code
        self.message = message
        super().__init__(f"Graph API Error ({status_code}): {error_code} - {message}")


class GraphClient:
    _ALLOWED_REDIRECT_DOMAINS = (".microsoft.com", ".sharepoint.com", ".office.com", ".office365.com")

    def __init__(
        self,
        base_url: str,
        token_provider: Callable[[], str],
        cb: CircuitBreaker | None = None,
    ):
        self.base_url = base_url.rstrip("/")
        self._token_provider = token_provider
        self._cb = cb
        self._client: httpx.AsyncClient | None = None

    async def __aenter__(self) -> "GraphClient":
        self._client = httpx.AsyncClient(
            headers={"Content-Type": "application/json"},
            timeout=30.0,
        )
        return self

    async def __aexit__(self, exc_type, exc_val, exc_tb) -> None:
        if self._client:
            await self._client.aclose()
            self._client = None

    @staticmethod
    def calendar_headers(timezone: str) -> dict[str, str]:
        return {"Prefer": f'outlook.timezone="{timezone}"'}

    async def get(
        self,
        endpoint: str,
        params: dict[str, Any] | None = None,
        headers: dict[str, str] | None = None,
    ) -> dict:
        return await self._request("GET", endpoint, params=params, headers=headers)

    async def post(
        self,
        endpoint: str,
        json_data: dict[str, Any],
        headers: dict[str, str] | None = None,
    ) -> dict:
        return await self._request("POST", endpoint, json_data=json_data, headers=headers)

    async def get_all(
        self,
        endpoint: str,
        params: dict[str, Any] | None = None,
        headers: dict[str, str] | None = None,
        max_pages: int = 5,
    ) -> tuple[list[dict], bool]:
        """Fetch all items from a paginated endpoint, following @odata.nextLink.

        Returns (items, has_more) where has_more indicates truncation.
        """
        items: list[dict] = []
        current_endpoint: str | None = endpoint
        current_params = params
        pages_fetched = 0

        while current_endpoint and pages_fetched < max_pages:
            page = await self.get(current_endpoint, params=current_params, headers=headers)
            items.extend(page.get("value", []))
            pages_fetched += 1

            next_link = page.get("@odata.nextLink")
            if next_link:
                if next_link.startswith(self.base_url):
                    current_endpoint = next_link[len(self.base_url):]
                else:
                    current_endpoint = next_link
                current_params = None
            else:
                current_endpoint = None

        has_more = current_endpoint is not None
        return items, has_more

    async def get_content(
        self,
        endpoint: str,
        headers: dict[str, str] | None = None,
    ) -> bytes:
        """Download raw file content. Validates redirect targets against allowed domains."""
        if not self._client:
            raise RuntimeError("Client not initialized. Use async context manager.")

        auth_headers = {"Authorization": f"Bearer {self._token_provider()}"}
        merged = {**auth_headers, **(headers or {})}
        url = f"{self.base_url}{endpoint}"

        response = await self._client.get(url, headers=merged, follow_redirects=False)

        if response.status_code in (301, 302, 307, 308):
            redirect_url = response.headers.get("location", "")
            if not self._is_allowed_redirect(redirect_url):
                raise GraphAPIError(
                    403, "RedirectBlocked",
                    f"Redirect to disallowed domain: {redirect_url[:100]}"
                )
            response = await self._client.get(redirect_url, headers=merged, follow_redirects=False)

        self._log_request("GET", endpoint, response)

        if response.status_code >= 400:
            raise self._parse_error(response)

        return response.content

    @staticmethod
    def _is_allowed_redirect(url: str) -> bool:
        from urllib.parse import urlparse
        try:
            hostname = urlparse(url).hostname or ""
            return any(hostname.endswith(domain) for domain in GraphClient._ALLOWED_REDIRECT_DOMAINS)
        except Exception:
            return False

    @staticmethod
    def _parse_error(response: httpx.Response) -> GraphAPIError:
        try:
            error_data = response.json()
            error = error_data.get("error", {})
            error_code = error.get("code", "UnknownError")
            message = error.get("message", response.text)
        except Exception:
            error_code = "UnknownError"
            message = response.text or f"HTTP {response.status_code}"
        return GraphAPIError(response.status_code, error_code, message)

    async def _request(
        self,
        method: str,
        endpoint: str,
        params: dict[str, Any] | None = None,
        json_data: dict[str, Any] | None = None,
        headers: dict[str, str] | None = None,
    ) -> dict:
        if not self._client:
            raise RuntimeError("Client not initialized. Use async context manager.")

        auth_headers = {"Authorization": f"Bearer {self._token_provider()}"}
        merged = {**auth_headers, **(headers or {})}
        url = f"{self.base_url}{endpoint}"

        async def _do_request():
            if method == "GET":
                return await self._client.get(url, params=params, headers=merged)
            elif method == "POST":
                return await self._client.post(url, json=json_data, headers=merged)
            raise ValueError(f"Unsupported method: {method}")

        if self._cb is not None:
            response = await self._cb.call(_do_request)
        else:
            response = await _do_request()

        if response.status_code in (429, 503):
            retry_after = min(int(response.headers.get("Retry-After", "1")), 10)
            _logger.warning("graph_api retry_after=%d status=%d endpoint=%s", retry_after, response.status_code, endpoint)
            await asyncio.sleep(retry_after)
            if self._cb is not None:
                response = await self._cb.call(_do_request)
            else:
                response = await _do_request()

        self._log_request(method, endpoint, response)
        return self._handle_response(response)

    def _log_request(self, method: str, endpoint: str, response: httpx.Response) -> None:
        size = len(response.content) if response.content else 0
        _logger.info(
            "graph_api method=%s endpoint=%s status=%d bytes=%d",
            method, endpoint, response.status_code, size,
        )

    def _handle_response(self, response: httpx.Response) -> dict:
        if response.status_code >= 400:
            raise self._parse_error(response)

        if response.status_code == 204 or not response.content:
            return {}

        return response.json()
