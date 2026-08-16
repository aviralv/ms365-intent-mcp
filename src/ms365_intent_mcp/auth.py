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
        # SharePoint-audience tokens, cached per host. Azure AD tokens are
        # audience-bound: a token minted for `{tenant}-my.sharepoint.com` is
        # rejected (401) by `{tenant}.sharepoint.com` and vice versa, even
        # with identical scopes. So we cache one token per resource host.
        self._spo_tokens: dict[str, tuple[str, float]] = {}

    def ensure_authenticated(self) -> None:
        token = self._try_refresh()
        if token:
            return
        raise AuthenticationError("No valid token. Run: ms365-intent-mcp auth")

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

    def get_sharepoint_token(self, host: str) -> str:
        """Mint (or return cached) a SharePoint-audience token for ``host``.

        SharePoint Vroom endpoints reject Graph-audience tokens with 401 — the
        ``aud`` claim must match the resource. This does the same refresh-token
        exchange as ``_try_refresh`` but with ``scope=https://{host}/.default``,
        yielding a token whose audience is SharePoint Online. Same delegated
        grant, same ``refresh_token``, different audience — NOT a new consent.

        Tokens are cached per host: ``{tenant}-my.sharepoint.com`` (personal-
        drive recordings) and ``{tenant}.sharepoint.com`` (team-channel-site
        recordings) are distinct audiences and must be minted separately.

        Raises ``AuthenticationError`` when no refresh token is available or the
        exchange fails, or ``ValueError`` if ``host`` is not a SharePoint host.
        """
        # Defense-in-depth: `host` is interpolated into the token scope, so a
        # caller passing an untrusted host would mint a token for an arbitrary
        # resource. VroomClient already validates before calling, but guard
        # here too so this public method is safe on its own.
        if not (host and host.endswith(".sharepoint.com")):
            raise ValueError(f"Refusing to mint a token for non-SharePoint host: {host!r}")

        cached = self._spo_tokens.get(host)
        if cached and time.time() < cached[1]:
            return cached[0]

        refresh_token = self._read_refresh_token()
        if not refresh_token:
            raise AuthenticationError("No valid token available. Run: ms365-intent-mcp auth")

        data = {
            "client_id": self.config.client_id,
            "grant_type": "refresh_token",
            "refresh_token": refresh_token,
            "scope": f"https://{host}/.default",
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
            _logger.warning("SharePoint token exchange failed for %s: %s", host, e)
            raise AuthenticationError(f"Failed to acquire SharePoint token for {host}") from e

        access_token = result.get("access_token")
        if not access_token:
            raise AuthenticationError(f"No access_token in SharePoint token response for {host}")
        expires_at = time.time() + result.get("expires_in", 3600) - _EXPIRY_BUFFER_SECONDS
        self._spo_tokens[host] = (access_token, expires_at)
        return access_token

    def _read_refresh_token(self) -> str | None:
        """Read the refresh token from the token file. Shared by Graph refresh
        and per-host SharePoint token exchange."""
        if not self.config.token_path.exists():
            _logger.warning("Token file not found: %s", self.config.token_path)
            return None
        try:
            token_data = json.loads(self.config.token_path.read_text())
        except (json.JSONDecodeError, OSError) as e:
            _logger.warning("Failed to read token file: %s", e)
            return None
        return token_data.get("refresh_token")

    def _try_refresh(self) -> str | None:
        refresh_token = self._read_refresh_token()
        if not refresh_token:
            _logger.warning("No refresh_token available")
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
