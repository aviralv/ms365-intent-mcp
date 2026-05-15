"""Token management — silent refresh only, no MSAL at runtime."""

import json
import logging
import time
import urllib.parse
import urllib.request

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
        if self._access_token:
            return self._access_token
        raise AuthenticationError("No valid token available. Run: ms365-intent-mcp auth")

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

        data = urllib.parse.urlencode({
            "client_id": self.config.client_id,
            "grant_type": "refresh_token",
            "refresh_token": refresh_token,
            "scope": " ".join(self.config.scopes),
        }).encode()

        req = urllib.request.Request(
            self.TOKEN_URL,
            data=data,
            headers={"Content-Type": "application/x-www-form-urlencoded"},
        )

        try:
            with urllib.request.urlopen(req, timeout=30) as resp:
                result = json.loads(resp.read().decode())
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
        self.config.token_path.write_text(json.dumps(token_data, indent=2))
        self.config.token_path.chmod(0o600)
