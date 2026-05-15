"""Permission registry — decodes JWT scp claim, provides scope checks."""

from __future__ import annotations

import base64
import json
import logging

_logger = logging.getLogger("ms365_intent_mcp")


class PermissionRegistry:
    def __init__(self, scopes: list[str]):
        self._scopes = set(scopes)

    @classmethod
    def from_token(cls, access_token: str) -> "PermissionRegistry":
        try:
            parts = access_token.split(".")
            if len(parts) != 3:
                _logger.warning("Token is not a valid JWT (expected 3 parts)")
                return cls([])
            payload_b64 = parts[1]
            padding = 4 - len(payload_b64) % 4
            payload_b64 += "=" * padding
            payload = json.loads(base64.urlsafe_b64decode(payload_b64))
            scp = payload.get("scp", "")
            scopes = [s for s in scp.split(" ") if s]
            _logger.info("permission_registry scopes_found=%d scopes=%s", len(scopes), scopes)
            return cls(scopes)
        except Exception as e:
            _logger.warning("Failed to decode token scopes: %s", e)
            return cls([])

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
