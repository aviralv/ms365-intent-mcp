"""Token management — silent refresh only, no MSAL at runtime."""

import json
import logging
import os
import tempfile
import time

import httpx

from .config import Config

_logger = logging.getLogger("ms365_intent_mcp")
_EXPIRY_BUFFER_SECONDS = 300


class AuthenticationError(Exception):
    pass


class TokenManager:
    TOKEN_URL = "https://login.microsoftonline.com/common/oauth2/v2.0/token"

    def __init__(self, config: Config):
        self.config = config
        self._access_token: str | None = None
        self._expires_at: float = 0.0

    def ensure_authenticated(self) -> None:
        token = self._try_refresh()
        if token:
            return
        raise AuthenticationError(
            "No valid token. Run: ms365-intent-mcp auth"
        )

    def get_access_token(self) -> str:
        if self._access_token and time.time() < self._expires_at:
            return self._access_token
        token = self._try_refresh()
        if token:
            return token
        raise AuthenticationError("No valid token available. Run: ms365-intent-mcp auth")

    def peek_access_token(self) -> str | None:
        """Return the cached token without triggering refresh. Safe to call from async contexts."""
        return self._access_token

    def _try_refresh(self) -> str | None:
        if not self.config.token_path.exists():
            _logger.warning("Token file not found: %s", self.config.token_path)
            return None

        try:
            token_data = json.loads(self.config.token_path.read_text())
        except (json.JSONDecodeError, OSError) as e:
            _logger.warning("Failed to read token file: %s", e)
            return None

        refresh_token = token_data.get("refresh_token")
        if not refresh_token:
            _logger.warning("No refresh_token in token file")
            return None

        data = {
            "client_id": self.config.client_id,
            "grant_type": "refresh_token",
            "refresh_token": refresh_token,
            "scope": " ".join(self.config.scopes),
        }

        try:
            resp = httpx.post(
                self.TOKEN_URL,
                data=data,
                headers={"Content-Type": "application/x-www-form-urlencoded"},
                timeout=30.0,
            )
            resp.raise_for_status()
            result = resp.json()
        except Exception as e:
            _logger.warning("Token refresh failed: %s", e)
            return None

        if "access_token" in result:
            self._access_token = result["access_token"]
            self._expires_at = time.time() + result.get("expires_in", 3600) - _EXPIRY_BUFFER_SECONDS
            self._save_tokens(result)
            return self._access_token

        return None

    def _save_tokens(self, response: dict) -> None:
        self.config.token_path.parent.mkdir(parents=True, exist_ok=True)
        token_data = {
            "access_token": response["access_token"],
            "refresh_token": response.get("refresh_token"),
            "expires_in": response.get("expires_in", 3600),
            "scope": response.get("scope", ""),
            "token_type": response.get("token_type", "Bearer"),
        }
        # Atomic write: temp file in the same directory + os.replace.
        # Two writers (e.g. parallel server instances) racing on a plain
        # write_text() can produce concatenated JSON ("Extra data: line N").
        # os.replace is atomic across files on the same filesystem, so a
        # reader sees either the old or the new file, never a partial one.
        parent = self.config.token_path.parent
        fd, tmp_path = tempfile.mkstemp(prefix=".token.", suffix=".tmp", dir=parent)
        try:
            with os.fdopen(fd, "w") as f:
                json.dump(token_data, f, indent=2)
            os.chmod(tmp_path, 0o600)
            os.replace(tmp_path, self.config.token_path)
        except Exception:
            try:
                os.unlink(tmp_path)
            except OSError:
                pass
            raise
