"""Unit tests for vroom.py host/redirect allowlist guards (security-critical).

These cover the SSRF boundary and the post-review hardening (issue #29 review
by GPT-5 + Gemini): host-suffix checks must not be bypassable, redirect
targets must be https, and non-SharePoint hosts must be refused.
"""

from __future__ import annotations

import pytest

from ms365_intent_mcp.vroom import VroomClient, VroomError, _is_allowed_host, _is_allowed_redirect

# ---------- _is_allowed_host: authenticated-request host allowlist ----------


def test_allowed_host_accepts_sharepoint_subdomains():
    assert _is_allowed_host("sap-my.sharepoint.com")
    assert _is_allowed_host("sap.sharepoint.com")
    assert _is_allowed_host("contoso.sharepoint.com")


def test_allowed_host_rejects_lookalike_prefix():
    """`evilsharepoint.com` must NOT match — the leading dot in the suffix
    prevents the substring bypass."""
    assert not _is_allowed_host("evilsharepoint.com")


def test_allowed_host_rejects_suffix_smuggling():
    """`foo.sharepoint.com.attacker.com` must NOT match."""
    assert not _is_allowed_host("foo.sharepoint.com.attacker.com")


def test_allowed_host_rejects_empty():
    assert not _is_allowed_host("")


# ---------- _is_allowed_redirect: CDN redirect allowlist + scheme ----------


def test_redirect_accepts_https_sharepoint_and_cdn():
    assert _is_allowed_redirect("https://foo.sharepoint.com/x")
    assert _is_allowed_redirect("https://cdn.sharepointonline.com/x")
    assert _is_allowed_redirect("https://x.sharepoint-df.com/y")


def test_redirect_rejects_plaintext_downgrade():
    """Even an allowed host must be rejected over http:// — no downgrade."""
    assert not _is_allowed_redirect("http://foo.sharepoint.com/x")


def test_redirect_rejects_foreign_host():
    assert not _is_allowed_redirect("https://evil.com/x")
    assert not _is_allowed_redirect("https://evilsharepoint.com/x")
    assert not _is_allowed_redirect("https://foo.sharepoint.com.attacker.com/x")


def test_redirect_rejects_non_http_scheme():
    assert not _is_allowed_redirect("file:///etc/passwd")
    assert not _is_allowed_redirect("ftp://foo.sharepoint.com/x")


def test_redirect_trailing_dot_fqdn():
    """A trailing-dot FQDN (`foo.sharepoint.com.`) in a Location header must not
    slip through. urlparse keeps the trailing dot in .hostname, so it won't end
    with the bare `.sharepoint.com` suffix — documenting the behavior."""
    assert not _is_allowed_redirect("https://foo.sharepoint.com./x")


# ---------- download_transcript_to_file: 3xx must not be a silent success ----------


class _FakeStreamResponse:
    """Async-context-manager stand-in for httpx's streaming response."""

    def __init__(self, status_code: int, *, location: str = "", chunks=(), body=b""):
        self.status_code = status_code
        self.headers = {"location": location} if location else {}
        self._chunks = chunks
        self._body = body

    async def __aenter__(self):
        return self

    async def __aexit__(self, *exc):
        return False

    async def aclose(self):
        return None

    async def aread(self):
        return self._body

    async def aiter_bytes(self):
        for c in self._chunks:
            yield c


class _FakeAsyncClient:
    """Minimal client whose .stream() returns queued fake responses in order."""

    def __init__(self, responses):
        self._responses = list(responses)
        self.calls = []

    def stream(self, method, url, headers=None, timeout=None):
        self.calls.append(url)
        return self._responses.pop(0)


@pytest.mark.asyncio
async def test_download_second_redirect_is_error_not_silent_zero(tmp_path):
    """Regression (issue #29 review): the CDN-redirect target returning a 3xx
    must raise, not fall through and write a 0-byte file that looks like
    success. First hop 302 → allowed CDN; CDN hop returns 302 again → error."""
    dest = tmp_path / "out.vtt"
    vc = VroomClient(lambda host: "tok")
    vc._client = _FakeAsyncClient(
        [
            _FakeStreamResponse(302, location="https://cdn.sharepointonline.com/x"),
            _FakeStreamResponse(302, location="https://cdn.sharepointonline.com/y"),
        ]
    )
    with pytest.raises(VroomError):
        await vc.download_transcript_to_file(
            "https://sap-my.sharepoint.com", "d", "i", "t", str(dest)
        )
    assert not dest.exists()  # partial/empty file cleaned up


@pytest.mark.asyncio
async def test_download_follows_one_redirect_to_cdn(tmp_path):
    """Happy redirect path: streamContent 302 → allowed CDN → 200 with bytes."""
    dest = tmp_path / "out.vtt"
    vc = VroomClient(lambda host: "tok")
    vc._client = _FakeAsyncClient(
        [
            _FakeStreamResponse(302, location="https://cdn.sharepointonline.com/x"),
            _FakeStreamResponse(200, chunks=[b"WEBVTT\n", b"<v A>hi</v>\n"]),
        ]
    )
    n = await vc.download_transcript_to_file(
        "https://sap-my.sharepoint.com", "d", "i", "t", str(dest)
    )
    assert n == len(b"WEBVTT\n<v A>hi</v>\n")
    assert dest.read_text().startswith("WEBVTT")


@pytest.mark.asyncio
async def test_download_disallowed_redirect_blocked(tmp_path):
    """A 302 to a non-SharePoint host must be blocked, not followed."""
    dest = tmp_path / "out.vtt"
    vc = VroomClient(lambda host: "tok")
    vc._client = _FakeAsyncClient([_FakeStreamResponse(302, location="https://evil.com/x")])
    with pytest.raises(VroomError):
        await vc.download_transcript_to_file(
            "https://sap-my.sharepoint.com", "d", "i", "t", str(dest)
        )
    assert not dest.exists()


# ---------- list_recordings_children: drive_id resolution ----------


class _FakeJsonResponse:
    """Stand-in for httpx's response for the non-streaming _get_json path."""

    def __init__(self, status_code: int, payload: dict):
        self.status_code = status_code
        self._payload = payload
        self.text = ""

    def json(self):
        return self._payload


class _FakeJsonClient:
    """Client whose .get() returns a queued fake JSON response."""

    def __init__(self, response):
        self._response = response
        self.calls = []

    async def get(self, url, headers=None):
        self.calls.append(url)
        return self._response


@pytest.mark.asyncio
async def test_list_recordings_children_drive_id_from_odata_context():
    """When @odata.context carries a real drive id in /drives/{id}/ form, use it."""
    payload = {
        "@odata.context": "https://sap-my.sharepoint.com/x/_api/v2.1/$metadata#/drives/b!REALID/items",
        "value": [{"id": "item1", "name": "A.mp4", "parentReference": {"driveId": "b!REALID"}}],
    }
    vc = VroomClient(lambda host: "tok")
    vc._client = _FakeJsonClient(_FakeJsonResponse(200, payload))
    items, drive_id = await vc.list_recordings_children("https://sap-my.sharepoint.com/personal/u")
    assert drive_id == "b!REALID"
    assert len(items) == 1


@pytest.mark.asyncio
async def test_list_recordings_children_drive_id_falls_back_to_item_parent_ref():
    """Regression (issue #31): Vroom /children returns @odata.context with the
    `drives('default')` alias, which the /drives/{id}/ regex does NOT match.
    The drive id must then fall back to the items' parentReference.driveId —
    otherwise every own-drive recording carries an empty drive_id and the
    download hop can't build a /drives/{id}/items/{id} URL."""
    payload = {
        "@odata.context": "https://sap-my.sharepoint.com/personal/u/_api/v2.1/$metadata#drives('default')/items('root')/children",
        "value": [
            {"id": "item1", "name": "A.mp4", "parentReference": {"driveId": "b!DERIVED"}},
            {"id": "item2", "name": "B.mp4", "parentReference": {"driveId": "b!DERIVED"}},
        ],
    }
    vc = VroomClient(lambda host: "tok")
    vc._client = _FakeJsonClient(_FakeJsonResponse(200, payload))
    items, drive_id = await vc.list_recordings_children("https://sap-my.sharepoint.com/personal/u")
    assert drive_id == "b!DERIVED"
    assert len(items) == 2
