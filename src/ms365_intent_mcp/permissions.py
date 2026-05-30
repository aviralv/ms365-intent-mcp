"""Permission registry — decodes JWT scp claim, provides scope checks."""

from __future__ import annotations

import base64
import json
import logging
from collections.abc import Callable

_logger = logging.getLogger("ms365_intent_mcp")


class PermissionRegistry:
    def __init__(self, scopes: list[str]):
        self._scopes = set(scopes)

    @classmethod
    def from_token(cls, access_token: str) -> "PermissionRegistry":
        scopes = cls._decode_scopes(access_token)
        return cls(scopes)

    @classmethod
    def from_token_provider(cls, token_provider: Callable[[], str | None]) -> "LazyPermissionRegistry":
        return LazyPermissionRegistry(token_provider)

    @staticmethod
    def _decode_scopes(access_token: str) -> list[str]:
        try:
            parts = access_token.split(".")
            if len(parts) != 3:
                _logger.warning("Token is not a valid JWT (expected 3 parts)")
                return []
            payload_b64 = parts[1]
            padding = (4 - len(payload_b64) % 4) % 4
            payload_b64 += "=" * padding
            payload = json.loads(base64.urlsafe_b64decode(payload_b64))
            scp = payload.get("scp", "")
            scopes = [s for s in scp.split(" ") if s]
            _logger.info("permission_registry scopes_found=%d scopes=%s", len(scopes), scopes)
            return scopes
        except Exception as e:
            _logger.warning("Failed to decode token scopes: %s", e)
            return []

    def has(self, scope: str) -> bool:
        return scope in self._scopes

    def has_any(self, scopes: list[str]) -> bool:
        return bool(self._scopes & set(scopes))

    def check(self, scope: str) -> str | None:
        if self.has(scope):
            return None
        return (
            f"ℹ️  Unavailable — {scope} scope not granted.\n"
            f"   To enable: ms365-intent-mcp auth --add-scope {scope}"
        )

    @property
    def granted(self) -> set[str]:
        return self._scopes.copy()


class LazyPermissionRegistry(PermissionRegistry):
    """Re-derives scopes from the current token when it changes.

    The token_provider must NOT trigger I/O — pass a peek-style function that
    returns the cached token. Refresh is the GraphClient's job (running in a
    thread); this registry only re-decodes when it observes a new token string.
    """

    def __init__(self, token_provider: Callable[[], str | None]):
        self._token_provider = token_provider
        self._last_token: str | None = None
        super().__init__([])

    def _refresh_if_needed(self) -> None:
        token = self._token_provider()
        if token and token != self._last_token:
            self._last_token = token
            self._scopes = set(self._decode_scopes(token))

    def has(self, scope: str) -> bool:
        self._refresh_if_needed()
        return super().has(scope)

    def has_any(self, scopes: list[str]) -> bool:
        self._refresh_if_needed()
        return super().has_any(scopes)

    def check(self, scope: str) -> str | None:
        self._refresh_if_needed()
        return super().check(scope)

    @property
    def granted(self) -> set[str]:
        self._refresh_if_needed()
        return super().granted
