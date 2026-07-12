"""Async HTTP client for SharePoint Vroom (2.1) transcript endpoints.

Sits *beside* ``GraphClient`` — deliberately NOT folded into it. Vroom lives on
tenant-dynamic ``*.sharepoint.com`` hosts and needs a SharePoint-audience
token, whereas ``GraphClient`` is pinned to ``graph.microsoft.com`` with a
Graph-audience token. Widening ``GraphClient``'s allowlist to a tenant-dynamic
wildcard would erode its SSRF boundary, so this is a separate narrow client.

Ported from ferret-transcripts ``core.py`` (sync ``requests`` → async
``httpx``). The pure parsing layer lives in ``transcripts.py``; this module is
the I/O half.

Security posture:
  - Authenticated requests only go to ``*.sharepoint.com`` hosts.
  - ``streamContent`` may 302 to a pre-signed CDN URL. We follow such redirects
    but do NOT forward the bearer token to non-SharePoint hosts (the CDN URL is
    already signed), and we never log redirect targets (they can carry SAS
    tokens).
  - A 403 on a single recording (cross-organizer, not-yet-opened-in-Teams) is a
    per-item permission gap, surfaced as a typed error — distinct from a 401
    audience mismatch, which is a token-layer bug.
"""

from __future__ import annotations

import asyncio
import logging
import os
import re
from collections.abc import Callable
from urllib.parse import quote, urlparse

import httpx

_logger = logging.getLogger("ms365_intent_mcp")

VROOM_BASE = "/_api/v2.1"

# Authenticated requests are restricted to these host suffixes.
_ALLOWED_HOST_SUFFIXES = (".sharepoint.com",)

# streamContent 302s may target these CDN hosts with a pre-signed URL. We
# follow the redirect (without the bearer token) but reject anything else.
_ALLOWED_REDIRECT_SUFFIXES = (
    ".sharepoint.com",
    ".sharepointonline.com",
    ".sharepoint-df.com",
)

# Explicit timeouts — httpx defaults to 5s, which a slow streamContent
# generation blows past. ferret's requests calls used 15–30s.
_JSON_TIMEOUT = httpx.Timeout(20.0)
_STREAM_TIMEOUT = httpx.Timeout(30.0, read=60.0)


class VroomError(Exception):
    """Raised on a non-recoverable Vroom API failure.

    ``status_code`` is the HTTP status (0 for transport-level errors). A 403 is
    a per-recording permission gap; a 401 is a token-audience problem.
    """

    def __init__(self, status_code: int, message: str):
        self.status_code = status_code
        self.message = message
        super().__init__(f"Vroom API Error ({status_code}): {message}")


def _host_of(url: str) -> str:
    return urlparse(url).hostname or ""


def _is_allowed_host(host: str) -> bool:
    return any(host.endswith(suffix) for suffix in _ALLOWED_HOST_SUFFIXES)


def _is_allowed_redirect(url: str) -> bool:
    host = _host_of(url)
    return any(host.endswith(suffix) for suffix in _ALLOWED_REDIRECT_SUFFIXES)


class VroomClient:
    """Narrow async client for SharePoint Vroom transcript operations.

    ``token_provider`` is called with a SharePoint host (e.g.
    ``sap-my.sharepoint.com``) and must return a bearer token whose audience
    matches that host — see ``TokenManager.get_sharepoint_token``.
    """

    def __init__(self, token_provider: Callable[[str], str]):
        self._token_provider = token_provider
        self._client: httpx.AsyncClient | None = None

    async def __aenter__(self) -> VroomClient:
        self._client = httpx.AsyncClient(
            headers={"accept": "application/json"},
            timeout=_JSON_TIMEOUT,
            follow_redirects=False,
        )
        return self

    async def __aexit__(self, exc_type, exc_val, exc_tb) -> None:
        if self._client:
            await self._client.aclose()
            self._client = None

    async def _token_for(self, url: str) -> str:
        host = _host_of(url)
        if not _is_allowed_host(host):
            raise VroomError(
                403, f"Refusing to send authenticated request to non-SharePoint host: {host}"
            )
        return await asyncio.to_thread(self._token_provider, host)

    async def _get_json(self, url: str) -> dict:
        if not self._client:
            raise RuntimeError("Client not initialized. Use async context manager.")
        token = await self._token_for(url)
        resp = await self._client.get(url, headers={"authorization": f"Bearer {token}"})
        self._log(url, resp.status_code)
        if resp.status_code >= 400:
            raise VroomError(resp.status_code, resp.text[:200])
        return resp.json()

    def _log(self, url: str, status: int) -> None:
        # Log only the API path, never the full URL — streamContent URLs and
        # redirect targets can carry SAS/presigned tokens.
        parsed = urlparse(url)
        _logger.info("vroom_api host=%s path=%s status=%d", parsed.hostname, parsed.path, status)

    # ------------------------------------------------------------------
    # Discovery / listing
    # ------------------------------------------------------------------

    async def list_recordings_children(
        self, site_root: str, top: int = 100
    ) -> tuple[list[dict], str]:
        """List items in the own-drive Recordings folder via Vroom /children.

        Returns ``(items, drive_id)``. ``drive_id`` is parsed from the
        ``@odata.context`` so callers can build direct drive/item URLs.
        Oversamples (``$top``) because Vroom 2.1 /children doesn't reliably
        honor ``$orderby`` — the caller runs pair-aware selection over the full
        window before slicing to a limit.
        """
        url = f"{site_root}{VROOM_BASE}/drive/root:/Recordings:/children?$top={top}"
        payload = await self._get_json(url)
        items = payload.get("value", [])
        drive_match = re.search(r"/drives/([^/]+)/", payload.get("@odata.context", ""))
        drive_id = drive_match.group(1) if drive_match else ""
        return items, drive_id

    async def resolve_item_by_filename(
        self, site_root: str, filename: str
    ) -> tuple[str, str]:
        """Resolve a Recordings-folder filename to ``(drive_id, item_id)`` via
        Vroom path-addressing. Returns ``("", "")`` when not found (404)."""
        encoded = quote(filename, safe="")
        url = f"{site_root}{VROOM_BASE}/drive/root:/Recordings/{encoded}:"
        try:
            payload = await self._get_json(url)
        except VroomError as exc:
            if exc.status_code == 404:
                return "", ""
            raise
        drive_id = (payload.get("parentReference") or {}).get("driveId", "")
        item_id = payload.get("id", "")
        return drive_id, item_id

    async def list_transcripts(
        self, site_root: str, drive_id: str, item_id: str
    ) -> list[dict]:
        """List transcript media for a recording. Returns the raw transcript
        dicts (each has an ``id``). Empty list when the recording has no
        transcript media (404 is treated as empty, not an error)."""
        url = (
            f"{site_root}{VROOM_BASE}/drives/{drive_id}/items/{item_id}"
            f"/media/transcripts"
        )
        try:
            payload = await self._get_json(url)
        except VroomError as exc:
            if exc.status_code == 404:
                return []
            raise
        return payload.get("value", [])

    # ------------------------------------------------------------------
    # Download
    # ------------------------------------------------------------------

    async def download_transcript_to_file(
        self,
        site_root: str,
        drive_id: str,
        item_id: str,
        transcript_id: str,
        dest_path: str,
    ) -> int:
        """Stream a transcript VTT to ``dest_path``. Returns bytes written.

        Cleans up a partial file if the download errors or is cancelled
        mid-stream. Handles a single 302 to a pre-signed CDN URL: the redirect
        is followed *without* the bearer token (the CDN URL is already signed),
        and its target is validated against the redirect allowlist.
        """
        if not self._client:
            raise RuntimeError("Client not initialized. Use async context manager.")

        url = (
            f"{site_root}{VROOM_BASE}/drives/{drive_id}/items/{item_id}"
            f"/media/transcripts/{transcript_id}"
            f"/streamContent?is=1&applymediaedits=false"
        )
        token = await self._token_for(url)
        headers: dict[str, str] = {"authorization": f"Bearer {token}"}

        bytes_written = 0
        try:
            async with self._client.stream(
                "GET", url, headers=headers, timeout=_STREAM_TIMEOUT
            ) as resp:
                # Follow one redirect hop to a pre-signed CDN URL. Drop the
                # bearer token — the redirect target is signed, and forwarding
                # the token to a CDN host would leak it.
                if resp.status_code in (301, 302, 307, 308):
                    location = resp.headers.get("location", "")
                    if not _is_allowed_redirect(location):
                        raise VroomError(403, "Redirect to disallowed host")
                    await resp.aclose()
                    return await self._stream_to_file(location, {}, dest_path)

                self._log(url, resp.status_code)
                if resp.status_code >= 400:
                    body = await resp.aread()
                    raise VroomError(resp.status_code, body[:200].decode("utf-8", "replace"))

                with open(dest_path, "wb") as fh:
                    async for chunk in resp.aiter_bytes():
                        fh.write(chunk)
                        bytes_written += len(chunk)
        except BaseException:
            _cleanup_partial(dest_path)
            raise
        return bytes_written

    async def _stream_to_file(
        self, url: str, headers: dict[str, str], dest_path: str
    ) -> int:
        """Stream a (pre-signed) URL to disk. Used for the CDN-redirect hop."""
        if not self._client:
            raise RuntimeError("Client not initialized. Use async context manager.")
        bytes_written = 0
        try:
            async with self._client.stream(
                "GET", url, headers=headers, timeout=_STREAM_TIMEOUT
            ) as resp:
                self._log(url, resp.status_code)
                if resp.status_code >= 400:
                    body = await resp.aread()
                    raise VroomError(resp.status_code, body[:200].decode("utf-8", "replace"))
                with open(dest_path, "wb") as fh:
                    async for chunk in resp.aiter_bytes():
                        fh.write(chunk)
                        bytes_written += len(chunk)
        except BaseException:
            _cleanup_partial(dest_path)
            raise
        return bytes_written


def _cleanup_partial(path: str) -> None:
    """Remove a partially-written file, ignoring errors."""
    try:
        if os.path.exists(path):
            os.unlink(path)
    except OSError:
        pass
