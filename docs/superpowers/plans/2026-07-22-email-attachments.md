# Email Attachment Extraction on `resolve` — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Let `resolve` enumerate an email's attachments (inline CID images + regular files) as metadata, and download their bytes to disk when `output_dir` is supplied.

**Architecture:** A new `composers/attachments.py` module owns enumeration (`GET /messages/{id}/attachments`, defensively paginated) and download (decode inline `contentBytes`, else stream `.../{id}/$value` via `GraphClient.get_content`, with collision-safe filenames + size caps). `composers/resolve.py`'s email branch calls it conditionally (gated on `hasAttachments`/`cid:`). New `AttachmentInfo` pydantic model extends `EmailContent`; `output_dir` extends `ResolvePayload`. The formatter renders an attachments section.

**Tech Stack:** Python 3.11+, FastMCP, httpx (async), pydantic v2, pytest + pytest-asyncio.

## Global Constraints

- **Version bump:** `pyproject.toml` `version` `1.5.0` → `1.6.0` (minor). Part of the change, not a separate step.
- **9-tool ceiling:** no new tool. Extend `resolve` only.
- **No raw IDs in markdown responses** — names, subjects, links, local paths only (attachment_id lives in structured data, never in `rendered_markdown`).
- **Markdown output only** (no JSON format branch).
- **Partial success:** attachment enumeration/download failures must never fail the whole `resolve`; the email body still renders.
- **`GraphClient` allowlist unchanged** — `$value` is a `graph.microsoft.com` path, so it uses the existing `GraphClient`; do NOT touch `VroomClient` or the host allowlist.
- **Surgical changes:** every changed line traces to this task. Don't reformat adjacent code.
- Constants for caps (module-level in `attachments.py`): `MAX_ATTACHMENT_BYTES = 100 * 1024 * 1024`, `MAX_TOTAL_ATTACHMENT_BYTES = 250 * 1024 * 1024`.

---

## File Structure

- **Create** `src/ms365_intent_mcp/composers/attachments.py` — enumeration + download helpers, filename safety, size caps. One responsibility: turning a message endpoint into attachment metadata and (optionally) files on disk.
- **Create** `tests/test_composers_attachments.py` — unit tests for the helpers.
- **Modify** `src/ms365_intent_mcp/intent/resolve/schemas.py` — add `AttachmentInfo`, extend `EmailContent`, extend `ResolvePayload`.
- **Modify** `src/ms365_intent_mcp/composers/resolve.py` — email branch wiring: enumerate + download, thread `output_dir` through `compose_resolve` and `_fetch_resolved`; extend `_build_structured_data` email case.
- **Modify** `src/ms365_intent_mcp/formatters.py` — attachments section in `format_resolved_content_markdown` email branch.
- **Modify** `src/ms365_intent_mcp/intent/resolve/impl.py` — pass `payload.output_dir` into `compose_resolve`.
- **Modify** `tests/test_composers_resolve.py` — email-branch wiring tests (enumeration gating, download).
- **Modify** `tests/intent/test_resolve_impl.py` — update the mocked `compose_resolve` signature to accept `output_dir`; assert `output_dir` round-trip.
- **Modify** `pyproject.toml` — version bump.
- **Regenerate** `tests/intent/snapshots/schemas/*.json` — new schema fields.

---

## Task 1: Attachment classification + filename safety (pure helpers)

**Files:**
- Create: `src/ms365_intent_mcp/composers/attachments.py`
- Test: `tests/test_composers_attachments.py`

**Interfaces:**
- Produces:
  - `ATTACHMENT_FILE_TYPE = "#microsoft.graph.fileAttachment"` (and `_ITEM_TYPE`, `_REFERENCE_TYPE` constants)
  - `MAX_ATTACHMENT_BYTES`, `MAX_TOTAL_ATTACHMENT_BYTES` constants
  - `def classify_attachment(raw: dict) -> dict` — maps one raw Graph attachment dict to a metadata dict: `{"name", "content_type", "size", "is_inline", "cid", "attachment_id", "kind", "_content_bytes", "note", "local_path"}` where `kind ∈ {"file","item","reference"}`, `cid` is the bracket-stripped `contentId` (`""` if none), `_content_bytes` is the raw base64 string or `None`, `local_path` is `None`, `note` is `None`.
  - `def safe_filename(name: str, index: int, existing: set[str]) -> str` — sanitized, collision-suffixed, ≤255-char filename; falls back to `attachment_{index}` for empty/degenerate names.

- [ ] **Step 1: Write the failing tests**

```python
# tests/test_composers_attachments.py
import pytest

from ms365_intent_mcp.composers.attachments import (
    classify_attachment,
    safe_filename,
)


class TestClassifyAttachment:
    def test_inline_image_carries_cid_and_bytes(self):
        raw = {
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": "image001.png",
            "contentType": "image/png",
            "size": 1234,
            "isInline": True,
            "contentId": "image001.png@01DD.CD8E",
            "contentBytes": "aGVsbG8=",
            "id": "AAMkID1",
        }
        m = classify_attachment(raw)
        assert m["kind"] == "file"
        assert m["is_inline"] is True
        assert m["cid"] == "image001.png@01DD.CD8E"
        assert m["content_type"] == "image/png"
        assert m["attachment_id"] == "AAMkID1"
        assert m["_content_bytes"] == "aGVsbG8="
        assert m["local_path"] is None
        assert m["note"] is None

    def test_cid_angle_brackets_stripped(self):
        raw = {
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": "x.png", "size": 1, "isInline": True,
            "contentId": "<abc@xyz>", "contentBytes": "AA==", "id": "i",
        }
        assert classify_attachment(raw)["cid"] == "abc@xyz"

    def test_regular_file_has_empty_cid(self):
        raw = {
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": "report.pdf", "contentType": "application/pdf",
            "size": 5000, "isInline": False, "id": "i2",
            "contentBytes": "AA==",
        }
        m = classify_attachment(raw)
        assert m["kind"] == "file"
        assert m["is_inline"] is False
        assert m["cid"] == ""

    def test_item_attachment_flagged_not_downloadable(self):
        raw = {
            "@odata.type": "#microsoft.graph.itemAttachment",
            "name": "Fwd: hi", "size": 900, "id": "i3",
        }
        m = classify_attachment(raw)
        assert m["kind"] == "item"
        assert m["_content_bytes"] is None
        assert m["note"] and "embedded" in m["note"].lower()

    def test_reference_attachment_flagged_not_downloadable(self):
        raw = {
            "@odata.type": "#microsoft.graph.referenceAttachment",
            "name": "cloud.docx", "id": "i4",
        }
        m = classify_attachment(raw)
        assert m["kind"] == "reference"
        assert m["_content_bytes"] is None
        assert m["note"] and "link" in m["note"].lower()


class TestSafeFilename:
    def test_basic_name_kept(self):
        assert safe_filename("image001.png", 0, set()) == "image001.png"

    def test_collision_suffixed(self):
        existing = {"image001.png"}
        out = safe_filename("image001.png", 1, existing)
        assert out == "image001-2.png"

    def test_second_collision_increments(self):
        existing = {"image001.png", "image001-2.png"}
        assert safe_filename("image001.png", 2, existing) == "image001-3.png"

    def test_empty_name_falls_back(self):
        assert safe_filename("", 3, set()) == "attachment_3"

    def test_dotdot_name_falls_back(self):
        out = safe_filename("..", 4, set())
        assert out == "attachment_4"

    def test_traversal_chars_sanitized(self):
        out = safe_filename("../../etc/passwd", 0, set())
        assert "/" not in out and ".." not in out

    def test_overlong_name_truncated_preserving_ext(self):
        long = "a" * 300 + ".png"
        out = safe_filename(long, 0, set())
        assert len(out) <= 255
        assert out.endswith(".png")
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `uv run pytest tests/test_composers_attachments.py -v`
Expected: FAIL — `ModuleNotFoundError` / `ImportError: cannot import name 'classify_attachment'`.

- [ ] **Step 3: Write the module**

```python
# src/ms365_intent_mcp/composers/attachments.py
"""Email attachment enumeration + download for the resolve composer.

Turns a resolved message endpoint (e.g. ``/me/messages/{id}``) into attachment
metadata, and — when an output_dir is supplied — materializes fileAttachment
bytes to disk. Inline CID images and regular file attachments are the same
Graph resource type (``fileAttachment``); the only difference is ``isInline``
and whether ``contentId`` matches a ``cid:`` token in the body.

Graph facts (verified against Microsoft Learn 2026-07-22):
  * ``contentBytes`` is usually returned inline on a list/get, with NO documented
    read-time size cutoff (the 3-4 MB limit is upload-only). But it can be null
    for a fileAttachment (e.g. a $select excluding it), so a $value fallback
    exists — gated on "null despite fileAttachment + non-zero size", not size.
  * ``GET .../attachments/{id}/$value`` returns raw binary for a fileAttachment,
    MIME for an itemAttachment, and 405 for a referenceAttachment.
  * The attachments collection has no documented page size; follow
    ``@odata.nextLink`` defensively.
"""

from __future__ import annotations

import base64
import binascii
import re
from pathlib import Path

from ..graph import GraphAPIError, GraphClient

ATTACHMENT_FILE_TYPE = "#microsoft.graph.fileAttachment"
_ITEM_TYPE = "#microsoft.graph.itemAttachment"
_REFERENCE_TYPE = "#microsoft.graph.referenceAttachment"

MAX_ATTACHMENT_BYTES = 100 * 1024 * 1024
MAX_TOTAL_ATTACHMENT_BYTES = 250 * 1024 * 1024

_SANITIZE_RE = re.compile(r"[^A-Za-z0-9._-]+")
_MAX_FILENAME_LEN = 255


def classify_attachment(raw: dict) -> dict:
    """Map one raw Graph attachment dict to an internal metadata dict.

    ``kind`` is "file" | "item" | "reference". ``_content_bytes`` holds the raw
    base64 string for fileAttachments (or None). ``cid`` is the bracket-stripped
    contentId ("" when absent). ``note``/``local_path`` start None; the download
    step fills them.
    """
    odata_type = raw.get("@odata.type", "")
    name = raw.get("name") or ""
    size = raw.get("size") or 0
    attachment_id = raw.get("id") or ""

    if odata_type == _ITEM_TYPE:
        kind, content_bytes, note = "item", None, "embedded item — not a downloadable file"
    elif odata_type == _REFERENCE_TYPE:
        kind, content_bytes, note = "reference", None, "cloud link — not a downloadable file"
    else:
        kind, content_bytes, note = "file", raw.get("contentBytes"), None

    cid_raw = raw.get("contentId") or ""
    cid = cid_raw.strip().lstrip("<").rstrip(">") if cid_raw else ""

    return {
        "name": name,
        "content_type": raw.get("contentType") or "",
        "size": size,
        "is_inline": bool(raw.get("isInline")),
        "cid": cid,
        "attachment_id": attachment_id,
        "kind": kind,
        "_content_bytes": content_bytes,
        "note": note,
        "local_path": None,
    }


def safe_filename(name: str, index: int, existing: set[str]) -> str:
    """Return a filesystem-safe, collision-free filename ≤255 chars.

    Sanitizes to ``[A-Za-z0-9._-]``; falls back to ``attachment_{index}`` for
    empty/dot-only/degenerate names; truncates over-long names preserving the
    extension; suffixes ``-2``, ``-3`` … on collision against ``existing``.
    """
    cleaned = _SANITIZE_RE.sub("_", name).strip("_")
    if not cleaned or set(cleaned) <= {".", "_"}:
        cleaned = f"attachment_{index}"

    if len(cleaned) > _MAX_FILENAME_LEN:
        stem, dot, ext = cleaned.rpartition(".")
        if dot and len(ext) < 20:
            keep = _MAX_FILENAME_LEN - len(ext) - 1
            cleaned = f"{stem[:keep]}.{ext}"
        else:
            cleaned = cleaned[:_MAX_FILENAME_LEN]

    candidate = cleaned
    if candidate not in existing:
        return candidate
    stem, dot, ext = cleaned.rpartition(".")
    base, suffix = (stem, f".{ext}") if dot else (cleaned, "")
    n = 2
    while True:
        candidate = f"{base}-{n}{suffix}"
        if candidate not in existing:
            return candidate
        n += 1
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `uv run pytest tests/test_composers_attachments.py -v`
Expected: PASS (all classify + safe_filename tests).

- [ ] **Step 5: Commit**

```bash
git add src/ms365_intent_mcp/composers/attachments.py tests/test_composers_attachments.py
git commit -m "feat(#49): attachment classification + safe filename helpers"
```

---

## Task 2: Enumeration with defensive pagination

**Files:**
- Modify: `src/ms365_intent_mcp/composers/attachments.py`
- Test: `tests/test_composers_attachments.py`

**Interfaces:**
- Consumes: `classify_attachment` (Task 1); `GraphClient.get` (`async get(endpoint, params=None, headers=None) -> dict`), `GraphAPIError`.
- Produces:
  - `def body_has_cid(body_text: str) -> bool` — True if the body contains a `cid:` reference.
  - `async def enumerate_attachments(client: GraphClient, message_endpoint: str) -> tuple[list[dict], str | None]` — returns `(metadata_list, error_note)`. Fetches `{message_endpoint}/attachments?$top=100`, follows `@odata.nextLink` (max 5 pages). On `GraphAPIError`, returns `([], reason)`. `message_endpoint` is the resolved email endpoint, e.g. `/me/messages/{id}`.

- [ ] **Step 1: Write the failing tests**

```python
# append to tests/test_composers_attachments.py
from unittest.mock import AsyncMock

from ms365_intent_mcp.composers.attachments import (
    body_has_cid,
    enumerate_attachments,
)
from ms365_intent_mcp.graph import GraphAPIError


class TestBodyHasCid:
    def test_detects_cid_token(self):
        assert body_has_cid('see <img src="cid:image001@01DD">') is True

    def test_detects_bracketed_cid_text(self):
        assert body_has_cid("inline [cid:image001.png@01DD.CD8E]") is True

    def test_no_cid(self):
        assert body_has_cid("just plain text") is False


class TestEnumerateAttachments:
    @pytest.mark.asyncio
    async def test_single_page(self):
        client = AsyncMock()
        client.get = AsyncMock(return_value={
            "value": [
                {"@odata.type": "#microsoft.graph.fileAttachment",
                 "name": "a.png", "size": 10, "isInline": True,
                 "contentId": "a@1", "contentBytes": "AA==", "id": "i1"},
            ]
        })
        entries, err = await enumerate_attachments(client, "/me/messages/M1")
        assert err is None
        assert len(entries) == 1
        assert entries[0]["cid"] == "a@1"
        client.get.assert_awaited()  # first call hits /attachments

    @pytest.mark.asyncio
    async def test_follows_next_link(self):
        client = AsyncMock()
        page1 = {
            "value": [{"@odata.type": "#microsoft.graph.fileAttachment",
                       "name": "a.png", "size": 1, "id": "i1", "contentBytes": "AA=="}],
            "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/messages/M1/attachments?$skip=100",
        }
        page2 = {
            "value": [{"@odata.type": "#microsoft.graph.fileAttachment",
                       "name": "b.png", "size": 1, "id": "i2", "contentBytes": "AA=="}],
        }
        client.get = AsyncMock(side_effect=[page1, page2])
        entries, err = await enumerate_attachments(client, "/me/messages/M1")
        assert err is None
        assert [e["name"] for e in entries] == ["a.png", "b.png"]
        assert client.get.await_count == 2

    @pytest.mark.asyncio
    async def test_graph_error_returns_note(self):
        client = AsyncMock()
        client.get = AsyncMock(side_effect=GraphAPIError(403, "ErrorAccessDenied", "no"))
        entries, err = await enumerate_attachments(client, "/me/messages/M1")
        assert entries == []
        assert err and "403" in err
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `uv run pytest tests/test_composers_attachments.py -k "BodyHasCid or Enumerate" -v`
Expected: FAIL — `ImportError: cannot import name 'body_has_cid'`.

- [ ] **Step 3: Add the functions**

```python
# add imports at top of attachments.py (merge with existing)
from ._utils import _error_reason

# add near the constants
_CID_RE = re.compile(r"cid:", re.IGNORECASE)
_ENUM_MAX_PAGES = 5


def body_has_cid(body_text: str) -> bool:
    """True if the body references an inline attachment via a cid: token.

    Works on both HTML bodies (``src="cid:..."``) and the plain-text body Graph
    returns under the ``outlook.body-content-type="text"`` Prefer header, where
    inline refs survive as ``[cid:...]`` tokens.
    """
    return bool(body_text) and _CID_RE.search(body_text) is not None


async def enumerate_attachments(
    client: GraphClient, message_endpoint: str
) -> tuple[list[dict], str | None]:
    """List a message's attachments as classified metadata dicts.

    Returns ``(entries, error_note)``. Follows ``@odata.nextLink`` defensively
    (the attachments collection has no documented page size). A GraphAPIError
    on any page returns whatever was collected plus a reason note — enumeration
    never raises to the caller.
    """
    entries: list[dict] = []
    next_url: str | None = f"{message_endpoint}/attachments?$top=100"
    pages = 0
    while next_url and pages < _ENUM_MAX_PAGES:
        try:
            resp = await client.get(next_url)
        except GraphAPIError as exc:
            return entries, f"attachment enumeration failed: {_error_reason(exc)}"
        for raw in resp.get("value", []):
            entries.append(classify_attachment(raw))
        next_url = resp.get("@odata.nextLink") or None
        pages += 1
    return entries, None
```

Note: `GraphClient.get` accepts absolute `graph.microsoft.com` URLs (the `@odata.nextLink` case) — see `graph.py` `_request` host check. No change needed there.

- [ ] **Step 4: Run tests to verify they pass**

Run: `uv run pytest tests/test_composers_attachments.py -k "BodyHasCid or Enumerate" -v`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/ms365_intent_mcp/composers/attachments.py tests/test_composers_attachments.py
git commit -m "feat(#49): attachment enumeration with defensive pagination"
```

---

## Task 3: Download (inline decode + `$value` stream + caps + containment)

**Files:**
- Modify: `src/ms365_intent_mcp/composers/attachments.py`
- Test: `tests/test_composers_attachments.py`

**Interfaces:**
- Consumes: `safe_filename`, `MAX_ATTACHMENT_BYTES`, `MAX_TOTAL_ATTACHMENT_BYTES` (Task 1); `GraphClient.get_content` (`async get_content(endpoint, headers=None) -> bytes`); metadata dicts from `enumerate_attachments` (Task 2).
- Produces:
  - `async def download_attachments(client: GraphClient, message_endpoint: str, entries: list[dict], output_dir: str | None) -> None` — mutates each `file`-kind entry in place: sets `local_path` on success, or `note` on skip/failure. Non-file kinds are left with their classification note. Enforces per-attachment and per-request byte caps. Writes under a resolved `output_dir` (default `~/.cache/ms365-intent-mcp/attachments`), containment-guarded.

- [ ] **Step 1: Write the failing tests**

```python
# append to tests/test_composers_attachments.py
import base64 as _b64

from ms365_intent_mcp.composers.attachments import (
    download_attachments,
    MAX_ATTACHMENT_BYTES,
)


def _file_entry(name, cid="", content_bytes="AA==", size=1, aid="i", ct="image/png"):
    return {
        "name": name, "content_type": ct, "size": size, "is_inline": bool(cid),
        "cid": cid, "attachment_id": aid, "kind": "file",
        "_content_bytes": content_bytes, "note": None, "local_path": None,
    }


class TestDownloadAttachments:
    @pytest.mark.asyncio
    async def test_inline_bytes_written(self, tmp_path):
        client = AsyncMock()
        payload = _b64.b64encode(b"hello").decode()
        entries = [_file_entry("image001.png", cid="a@1", content_bytes=payload, size=5)]
        await download_attachments(client, "/me/messages/M1", entries, str(tmp_path))
        p = entries[0]["local_path"]
        assert p is not None
        assert Path(p).read_bytes() == b"hello"
        client.get_content.assert_not_awaited()  # inline, no $value hop

    @pytest.mark.asyncio
    async def test_value_fallback_when_bytes_null(self, tmp_path):
        client = AsyncMock()
        client.get_content = AsyncMock(return_value=b"streamed")
        entries = [_file_entry("big.bin", content_bytes=None, size=9, aid="AT9")]
        await download_attachments(client, "/me/messages/M1", entries, str(tmp_path))
        client.get_content.assert_awaited_once()
        called_endpoint = client.get_content.await_args.args[0]
        assert called_endpoint == "/me/messages/M1/attachments/AT9/$value"
        assert Path(entries[0]["local_path"]).read_bytes() == b"streamed"

    @pytest.mark.asyncio
    async def test_item_attachment_not_downloaded(self, tmp_path):
        client = AsyncMock()
        entries = [{"name": "Fwd", "content_type": "", "size": 1, "is_inline": False,
                    "cid": "", "attachment_id": "i", "kind": "item",
                    "_content_bytes": None, "note": "embedded item — not a downloadable file",
                    "local_path": None}]
        await download_attachments(client, "/me/messages/M1", entries, str(tmp_path))
        assert entries[0]["local_path"] is None
        client.get_content.assert_not_awaited()

    @pytest.mark.asyncio
    async def test_oversized_attachment_skipped(self, tmp_path):
        client = AsyncMock()
        entries = [_file_entry("huge.bin", content_bytes=None, size=MAX_ATTACHMENT_BYTES + 1, aid="X")]
        await download_attachments(client, "/me/messages/M1", entries, str(tmp_path))
        assert entries[0]["local_path"] is None
        assert "too large" in entries[0]["note"].lower()
        client.get_content.assert_not_awaited()

    @pytest.mark.asyncio
    async def test_total_cap_skips_remainder(self, tmp_path, monkeypatch):
        import ms365_intent_mcp.composers.attachments as att
        monkeypatch.setattr(att, "MAX_TOTAL_ATTACHMENT_BYTES", 10)
        client = AsyncMock()
        payload = _b64.b64encode(b"1234567").decode()  # 7 bytes
        entries = [
            _file_entry("a.bin", content_bytes=payload, size=7, aid="a"),
            _file_entry("b.bin", content_bytes=payload, size=7, aid="b"),
        ]
        await download_attachments(client, "/me/messages/M1", entries, str(tmp_path))
        assert entries[0]["local_path"] is not None
        assert entries[1]["local_path"] is None
        assert "cap" in entries[1]["note"].lower()

    @pytest.mark.asyncio
    async def test_collision_suffix_on_disk(self, tmp_path):
        client = AsyncMock()
        payload = _b64.b64encode(b"x").decode()
        entries = [
            _file_entry("image001.png", cid="a@1", content_bytes=payload, size=1, aid="a"),
            _file_entry("image001.png", cid="b@2", content_bytes=payload, size=1, aid="b"),
        ]
        await download_attachments(client, "/me/messages/M1", entries, str(tmp_path))
        names = {Path(e["local_path"]).name for e in entries}
        assert names == {"image001.png", "image001-2.png"}

    @pytest.mark.asyncio
    async def test_bad_base64_noted_not_raised(self, tmp_path):
        client = AsyncMock()
        entries = [_file_entry("x.png", content_bytes="!!!notbase64!!!", size=3, aid="a")]
        await download_attachments(client, "/me/messages/M1", entries, str(tmp_path))
        assert entries[0]["local_path"] is None
        assert entries[0]["note"]

    @pytest.mark.asyncio
    async def test_output_dir_is_existing_file_errors_gracefully(self, tmp_path):
        client = AsyncMock()
        f = tmp_path / "afile"
        f.write_text("x")
        payload = _b64.b64encode(b"x").decode()
        entries = [_file_entry("x.png", content_bytes=payload, size=1)]
        await download_attachments(client, "/me/messages/M1", entries, str(f))
        assert entries[0]["local_path"] is None
        assert entries[0]["note"]
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `uv run pytest tests/test_composers_attachments.py -k Download -v`
Expected: FAIL — `ImportError: cannot import name 'download_attachments'`.

- [ ] **Step 3: Add the download function**

```python
# add to attachments.py

def _resolve_output_base(output_dir: str | None) -> Path:
    base = (
        Path(output_dir).expanduser()
        if output_dir
        else Path.home() / ".cache" / "ms365-intent-mcp" / "attachments"
    ).resolve()
    base.mkdir(parents=True, exist_ok=True)
    return base


async def download_attachments(
    client: GraphClient,
    message_endpoint: str,
    entries: list[dict],
    output_dir: str | None,
) -> None:
    """Materialize fileAttachment bytes to disk; mutate entries in place.

    Each ``kind == "file"`` entry gets a ``local_path`` on success or a ``note``
    on skip/failure. Non-file kinds keep their classification note. Enforces the
    per-attachment and per-request byte caps. Never raises for a single
    attachment — failures are per-entry notes.
    """
    try:
        base = _resolve_output_base(output_dir)
    except (OSError, ValueError) as exc:
        for e in entries:
            if e["kind"] == "file":
                e["note"] = f"could not prepare output dir: {exc}"
        return

    used_names: set[str] = set()
    total = 0
    for index, e in enumerate(entries):
        if e["kind"] != "file":
            continue
        size = e.get("size") or 0
        if size > MAX_ATTACHMENT_BYTES:
            e["note"] = f"skipped — too large ({size} bytes > per-file cap)"
            continue
        if total + size > MAX_TOTAL_ATTACHMENT_BYTES:
            e["note"] = "skipped — per-request size cap reached"
            continue

        try:
            data = await _fetch_bytes(client, message_endpoint, e)
        except (GraphAPIError, binascii.Error, ValueError) as exc:
            e["note"] = f"download failed: {exc}"
            continue

        fname = safe_filename(e["name"], index, used_names)
        dest = (base / fname).resolve()
        if dest.parent != base:
            e["note"] = "skipped — resolved path escaped output dir"
            continue
        try:
            dest.write_bytes(data)
        except OSError as exc:
            e["note"] = f"write failed: {exc}"
            continue

        used_names.add(fname)
        e["local_path"] = str(dest)
        total += len(data)
        _note_extension_mismatch(e)


async def _fetch_bytes(client: GraphClient, message_endpoint: str, entry: dict) -> bytes:
    """Return the attachment's bytes: decode inline contentBytes if present,
    else stream the $value endpoint. Raises on decode/transport failure."""
    raw = entry.get("_content_bytes")
    if raw:
        return base64.b64decode(raw, validate=True)
    endpoint = f"{message_endpoint}/attachments/{entry['attachment_id']}/$value"
    return await client.get_content(endpoint)


def _note_extension_mismatch(entry: dict) -> None:
    """Informational note when the filename extension disagrees with the
    declared content-type major class (e.g. image bytes named .pdf)."""
    ct = (entry.get("content_type") or "").split("/", 1)[0].lower()
    name = entry.get("name") or ""
    ext = name.rpartition(".")[2].lower()
    _img_exts = {"png", "jpg", "jpeg", "gif", "bmp", "webp", "tif", "tiff"}
    if ct == "image" and ext and ext not in _img_exts:
        entry["note"] = f"content-type is image/* but filename ends .{ext}"
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `uv run pytest tests/test_composers_attachments.py -v`
Expected: PASS (all tasks 1–3 tests).

- [ ] **Step 5: Commit**

```bash
git add src/ms365_intent_mcp/composers/attachments.py tests/test_composers_attachments.py
git commit -m "feat(#49): attachment download with \$value fallback, size caps, containment"
```

---

## Task 4: Schema — `AttachmentInfo`, extend `EmailContent` + `ResolvePayload`

**Files:**
- Modify: `src/ms365_intent_mcp/intent/resolve/schemas.py`
- Test: `tests/intent/test_resolve_impl.py` (schema-construction check)

**Interfaces:**
- Produces:
  - `class AttachmentInfo(BaseModel)` with fields `name: str`, `content_type: str = ""`, `size: int = 0`, `is_inline: bool = False`, `cid: str = ""`, `attachment_id: str = ""`, `local_path: str | None = None`, `note: str | None = None`.
  - `EmailContent.attachments: list[AttachmentInfo] = []`
  - `ResolvePayload.output_dir: str | None = None`

- [ ] **Step 1: Write the failing test**

```python
# append to tests/intent/test_resolve_impl.py
from ms365_intent_mcp.intent.resolve.schemas import AttachmentInfo


class TestAttachmentSchema:
    def test_email_content_accepts_attachments(self):
        from ms365_intent_mcp.intent.resolve.schemas import EmailContent
        ec = EmailContent(
            kind="email", subject="s", sender="a@b.com", body="",
            attachments=[AttachmentInfo(name="x.png", cid="a@1", is_inline=True,
                                        size=10, attachment_id="i", local_path="/tmp/x.png")],
        )
        assert ec.attachments[0].name == "x.png"
        assert ec.attachments[0].local_path == "/tmp/x.png"

    def test_email_content_attachments_default_empty(self):
        from ms365_intent_mcp.intent.resolve.schemas import EmailContent
        ec = EmailContent(kind="email", subject="s", sender="a@b.com", body="")
        assert ec.attachments == []

    def test_payload_accepts_output_dir(self):
        p = ResolvePayload.model_validate({
            "url": "https://outlook.office.com/mail/id/AA123",
            "output_dir": "/tmp/out",
        })
        assert p.output_dir == "/tmp/out"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `uv run pytest tests/intent/test_resolve_impl.py -k Attachment -v`
Expected: FAIL — `ImportError: cannot import name 'AttachmentInfo'`.

- [ ] **Step 3: Edit the schema**

In `src/ms365_intent_mcp/intent/resolve/schemas.py`, add the model above `EmailContent`:

```python
class AttachmentInfo(BaseModel):
    """One email attachment — metadata always; local_path only after download."""

    model_config = ConfigDict(extra="forbid")
    name: str
    content_type: str = ""
    size: int = 0
    is_inline: bool = False
    cid: str = ""
    attachment_id: str = ""
    local_path: str | None = None
    note: str | None = None
```

Extend `EmailContent` (add the field after `body`):

```python
class EmailContent(BaseModel):
    """Content for a resolved email message."""

    model_config = ConfigDict(extra="forbid")
    kind: Literal["email"]
    subject: str
    sender: str
    body: str
    attachments: list[AttachmentInfo] = []
```

Extend `ResolvePayload`:

```python
class ResolvePayload(BaseModel):
    """Resolve any M365 URL."""

    model_config = ConfigDict(extra="forbid")
    url: HttpUrl
    output_dir: str | None = Field(
        default=None,
        description=(
            "For email URLs: download the email's attachments (inline images + "
            "regular files) to this directory and return each file's local path. "
            "Metadata is always returned; passing output_dir triggers the byte "
            "download. Defaults to ~/.cache/ms365-intent-mcp/attachments when "
            "download is requested. No-op for non-email URLs."
        ),
    )
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `uv run pytest tests/intent/test_resolve_impl.py -k Attachment -v`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/ms365_intent_mcp/intent/resolve/schemas.py tests/intent/test_resolve_impl.py
git commit -m "feat(#49): AttachmentInfo schema + output_dir on ResolvePayload"
```

---

## Task 5: Wire enumeration + download into the resolve composer

**Files:**
- Modify: `src/ms365_intent_mcp/composers/resolve.py`
- Test: `tests/test_composers_resolve.py`

**Interfaces:**
- Consumes: `enumerate_attachments`, `download_attachments`, `body_has_cid` (Tasks 2–3).
- Produces (changed signatures):
  - `async def compose_resolve(client, permissions, url, output_dir: str | None = None) -> tuple[dict, str]`
  - `_fetch_resolved(client, resolved, output_dir: str | None = None)` — email branch requests `hasAttachments` in `$select`, and after fetching the message, enumerates (gated on `hasAttachments`/`cid:`) and optionally downloads. Returns the message dict augmented with `_attachments` (list of metadata dicts) and `_attachments_error`.
  - `_build_structured_data` email case includes `attachments` (list of the response-safe dicts, `_content_bytes` stripped).

- [ ] **Step 1: Write the failing tests**

```python
# append to tests/test_composers_resolve.py

class TestResolveEmailAttachments:
    @pytest.mark.asyncio
    async def test_enumerates_when_has_attachments(self, full_permissions):
        client = AsyncMock()
        client.get = AsyncMock(side_effect=[
            {  # message fetch
                "subject": "Bug report", "from": {"emailAddress": {"name": "Cust"}},
                "receivedDateTime": "2026-07-20T08:00:00Z",
                "body": {"contentType": "text", "content": "see [cid:img1@01DD]"},
                "hasAttachments": True,
            },
            {  # /attachments
                "value": [{"@odata.type": "#microsoft.graph.fileAttachment",
                           "name": "shot.png", "contentType": "image/png",
                           "size": 12, "isInline": True, "contentId": "img1@01DD",
                           "contentBytes": "AA==", "id": "AT1"}],
            },
        ])
        with patch("ms365_intent_mcp.composers.resolve.resolve_url") as mock_resolve:
            mock_resolve.return_value = ResolvedUrl(
                url_type="email", graph_endpoint="/me/messages/M1", required_scope="Mail.Read",
            )
            data, md = await compose_resolve(
                client=client, permissions=full_permissions,
                url="https://outlook.office365.com/mail/id/M1",
            )
        atts = data["data"]["attachments"]
        assert len(atts) == 1
        assert atts[0]["cid"] == "img1@01DD"
        assert atts[0]["local_path"] is None  # no output_dir → metadata only
        assert "_content_bytes" not in atts[0]  # internal field stripped
        assert "shot.png" in md

    @pytest.mark.asyncio
    async def test_no_extra_call_when_no_attachments_no_cid(self, full_permissions):
        client = AsyncMock()
        client.get = AsyncMock(return_value={
            "subject": "plain", "from": {"emailAddress": {"name": "A"}},
            "receivedDateTime": "2026-07-20T08:00:00Z",
            "body": {"contentType": "text", "content": "no images here"},
            "hasAttachments": False,
        })
        with patch("ms365_intent_mcp.composers.resolve.resolve_url") as mock_resolve:
            mock_resolve.return_value = ResolvedUrl(
                url_type="email", graph_endpoint="/me/messages/M2", required_scope="Mail.Read",
            )
            data, _ = await compose_resolve(
                client=client, permissions=full_permissions,
                url="https://outlook.office365.com/mail/id/M2",
            )
        assert client.get.await_count == 1  # message only, no /attachments
        assert data["data"]["attachments"] == []

    @pytest.mark.asyncio
    async def test_downloads_when_output_dir_given(self, full_permissions, tmp_path):
        import base64
        client = AsyncMock()
        client.get = AsyncMock(side_effect=[
            {"subject": "s", "from": {"emailAddress": {"name": "A"}},
             "receivedDateTime": "2026-07-20T08:00:00Z",
             "body": {"contentType": "text", "content": "x"}, "hasAttachments": True},
            {"value": [{"@odata.type": "#microsoft.graph.fileAttachment",
                        "name": "r.pdf", "contentType": "application/pdf", "size": 5,
                        "isInline": False, "contentBytes": base64.b64encode(b"hello").decode(),
                        "id": "AT2"}]},
        ])
        with patch("ms365_intent_mcp.composers.resolve.resolve_url") as mock_resolve:
            mock_resolve.return_value = ResolvedUrl(
                url_type="email", graph_endpoint="/me/messages/M3", required_scope="Mail.Read",
            )
            data, _ = await compose_resolve(
                client=client, permissions=full_permissions,
                url="https://outlook.office365.com/mail/id/M3",
                output_dir=str(tmp_path),
            )
        lp = data["data"]["attachments"][0]["local_path"]
        assert lp is not None
        from pathlib import Path
        assert Path(lp).read_bytes() == b"hello"

    @pytest.mark.asyncio
    async def test_enumeration_error_degrades(self, full_permissions):
        client = AsyncMock()
        client.get = AsyncMock(side_effect=[
            {"subject": "s", "from": {"emailAddress": {"name": "A"}},
             "receivedDateTime": "2026-07-20T08:00:00Z",
             "body": {"contentType": "text", "content": "x"}, "hasAttachments": True},
            GraphAPIError(403, "ErrorAccessDenied", "no"),
        ])
        with patch("ms365_intent_mcp.composers.resolve.resolve_url") as mock_resolve:
            mock_resolve.return_value = ResolvedUrl(
                url_type="email", graph_endpoint="/me/messages/M4", required_scope="Mail.Read",
            )
            data, md = await compose_resolve(
                client=client, permissions=full_permissions,
                url="https://outlook.office365.com/mail/id/M4",
            )
        assert data["data"]["attachments"] == []
        assert "s" in md  # body still renders (subject present)
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `uv run pytest tests/test_composers_resolve.py -k EmailAttachments -v`
Expected: FAIL — `compose_resolve() got an unexpected keyword argument 'output_dir'`.

- [ ] **Step 3: Edit the composer**

In `composers/resolve.py`:

Add imports near the top (with the other relative imports):

```python
from .attachments import body_has_cid, download_attachments, enumerate_attachments
```

Change `compose_resolve` signature and the `_fetch_resolved` call:

```python
async def compose_resolve(
    client: GraphClient,
    permissions: PermissionRegistry,
    url: str,
    output_dir: str | None = None,
) -> tuple[dict, str]:
```

and inside it:

```python
    try:
        data = await _fetch_resolved(client, resolved, output_dir)
```

Change `_fetch_resolved` signature:

```python
async def _fetch_resolved(
    client: GraphClient, resolved: ResolvedUrl, output_dir: str | None = None
) -> dict:
```

In the `email` branch, add `hasAttachments` to `$select` and enumerate/download after the fetch:

```python
    if url_type == "email":
        message = await client.get(
            endpoint,
            params={
                "$select": "subject,from,receivedDateTime,bodyPreview,body,toRecipients,ccRecipients,webLink,hasAttachments",
            },
            headers={"Prefer": 'outlook.body-content-type="text"'},
        )
        body_text = (message.get("body") or {}).get("content", "") or (message.get("bodyPreview") or "")
        if message.get("hasAttachments") or body_has_cid(body_text):
            entries, enum_error = await enumerate_attachments(client, endpoint)
            if entries and output_dir is not None:
                await download_attachments(client, endpoint, entries, output_dir)
            message["_attachments"] = entries
            message["_attachments_error"] = enum_error
        else:
            message["_attachments"] = []
            message["_attachments_error"] = None
        return message
```

In `_build_structured_data`, extend the email case to strip the internal field:

```python
    if url_type == "email":
        raw_atts = data.get("_attachments") or []
        attachments = [
            {k: v for k, v in a.items() if k not in ("_content_bytes", "kind")}
            for a in raw_atts
        ]
        return {
            "kind": "email",
            "subject": data.get("subject", ""),
            "sender": (data.get("from") or {}).get("emailAddress", {}).get("name", ""),
            "body": (data.get("body") or {}).get("content", "") or (data.get("bodyPreview") or ""),
            "attachments": attachments,
        }
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `uv run pytest tests/test_composers_resolve.py -v`
Expected: PASS (new attachment tests + all existing email/URL tests still green).

- [ ] **Step 5: Commit**

```bash
git add src/ms365_intent_mcp/composers/resolve.py tests/test_composers_resolve.py
git commit -m "feat(#49): wire attachment enumerate/download into resolve email branch"
```

---

## Task 6: Render attachments in the email formatter

**Files:**
- Modify: `src/ms365_intent_mcp/formatters.py`
- Test: `tests/test_formatters.py`

**Interfaces:**
- Consumes: the `_attachments` list on the email data dict (Task 5).
- Produces: `format_resolved_content_markdown` email branch appends an attachments section — one 📎 line per attachment: `name · human size · inline|file · saved path or note`.

- [ ] **Step 1: Write the failing test**

```python
# append to tests/test_formatters.py
from ms365_intent_mcp.formatters import format_resolved_content_markdown


class TestEmailAttachmentRendering:
    def test_renders_attachment_lines(self):
        data = {
            "subject": "Bug", "from": {"emailAddress": {"name": "Cust"}},
            "receivedDateTime": "2026-07-20T00:00:00Z",
            "body": {"contentType": "text", "content": "see attached"},
            "_attachments": [
                {"name": "shot.png", "content_type": "image/png", "size": 20480,
                 "is_inline": True, "cid": "a@1", "attachment_id": "i",
                 "kind": "file", "local_path": "/tmp/shot.png", "note": None},
                {"name": "notes.docx", "content_type": "", "size": 0,
                 "is_inline": False, "cid": "", "attachment_id": "j",
                 "kind": "item", "local_path": None,
                 "note": "embedded item — not a downloadable file"},
            ],
        }
        md = format_resolved_content_markdown("email", data)
        assert "shot.png" in md
        assert "/tmp/shot.png" in md
        assert "notes.docx" in md
        assert "not a downloadable file" in md

    def test_no_attachment_section_when_empty(self):
        data = {
            "subject": "plain", "from": {"emailAddress": {"name": "A"}},
            "receivedDateTime": "2026-07-20T00:00:00Z",
            "body": {"contentType": "text", "content": "hi"},
            "_attachments": [],
        }
        md = format_resolved_content_markdown("email", data)
        assert "📎" not in md
```

- [ ] **Step 2: Run test to verify it fails**

Run: `uv run pytest tests/test_formatters.py -k EmailAttachment -v`
Expected: FAIL (no attachment lines rendered / `📎` absent when present).

- [ ] **Step 3: Edit the formatter**

In `formatters.py`, in `format_resolved_content_markdown`'s `email` branch, after the body is appended (after the `if rendered:` block, before `return "\n".join(lines)`):

```python
        attachments = data.get("_attachments") or []
        if attachments:
            lines.append("")
            lines.append("**Attachments:**")
            for a in attachments:
                lines.append(_format_email_attachment_line(a))
        return "\n".join(lines)
```

Add the helper near `_format_attachment_links`:

```python
def _format_email_attachment_line(a: dict) -> str:
    """One 📎 line for an email attachment: name · size · inline/file · path|note."""
    parts = [a.get("name") or "(unnamed)"]
    size = a.get("size") or 0
    if size:
        parts.append(_human_size(size))
    parts.append("inline" if a.get("is_inline") else "file")
    if a.get("local_path"):
        parts.append(f"saved: `{a['local_path']}`")
    elif a.get("note"):
        parts.append(a["note"])
    return "📎 " + " · ".join(parts)


def _human_size(n: int) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024 or unit == "GB":
            return f"{n:.0f} {unit}" if unit == "B" else f"{n/1:.0f} {unit}" if False else f"{n:.1f} {unit}".replace(".0 ", " ")
        n /= 1024
    return f"{n:.1f} GB"
```

Note: keep `_human_size` simple — if the ternary above reads awkwardly during implementation, replace its body with a plain loop:

```python
def _human_size(n: float) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024 or unit == "GB":
            return f"{int(n)} {unit}" if unit == "B" else f"{n:.1f} {unit}"
        n /= 1024
```

Use the plain-loop version; it's the intended implementation.

- [ ] **Step 4: Run tests to verify they pass**

Run: `uv run pytest tests/test_formatters.py -k EmailAttachment -v`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/ms365_intent_mcp/formatters.py tests/test_formatters.py
git commit -m "feat(#49): render email attachments section in resolve markdown"
```

---

## Task 7: Thread `output_dir` through the intent impl

**Files:**
- Modify: `src/ms365_intent_mcp/intent/resolve/impl.py`
- Test: `tests/intent/test_resolve_impl.py`

**Interfaces:**
- Consumes: `compose_resolve(client, permissions, url, output_dir)` (Task 5); `ResolvePayload.output_dir` (Task 4).
- Produces: `_resolve_impl` passes `payload.output_dir` to `compose_resolve`.

- [ ] **Step 1: Update the existing mocked-signature test + add round-trip test**

The existing happy-path test mocks `_fake(client, permissions, url)` — it must accept `output_dir`. Update it and add a round-trip assertion:

```python
# in tests/intent/test_resolve_impl.py — update the existing _fake signatures
# to: async def _fake(client, permissions, url, output_dir=None):
# then add:

class TestResolveOutputDir:
    @pytest.mark.asyncio
    async def test_output_dir_threaded_to_composer(self, monkeypatch):
        ctx, _, _ = _mock_ctx()
        seen = {}

        async def _fake(client, permissions, url, output_dir=None):
            seen["output_dir"] = output_dir
            return (
                {"url": url, "kind": "email",
                 "data": {"kind": "email", "subject": "s", "sender": "a@b.com",
                          "body": "", "attachments": [
                              {"name": "x.png", "content_type": "image/png", "size": 5,
                               "is_inline": True, "cid": "a@1", "attachment_id": "i",
                               "local_path": "/tmp/out/x.png", "note": None}]}},
                "rendered",
            )

        monkeypatch.setattr(
            "ms365_intent_mcp.intent.resolve.impl.compose_resolve", _fake
        )
        payload = ResolvePayload.model_validate({
            "url": "https://outlook.office.com/mail/id/AA123",
            "output_dir": "/tmp/out",
        })
        response = await _resolve_impl(ctx, payload)
        assert seen["output_dir"] == "/tmp/out"
        assert response.data.attachments[0].local_path == "/tmp/out/x.png"
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `uv run pytest tests/intent/test_resolve_impl.py -k OutputDir -v`
Expected: FAIL — `compose_resolve` called without `output_dir` (seen value stays unset / TypeError from the updated `_fake`).

- [ ] **Step 3: Edit the impl**

In `intent/resolve/impl.py`, change the `compose_resolve` call:

```python
    data_dict, markdown = await compose_resolve(
        client, permissions, str(payload.url), payload.output_dir
    )
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `uv run pytest tests/intent/test_resolve_impl.py -v`
Expected: PASS (round-trip + all existing impl tests).

- [ ] **Step 5: Commit**

```bash
git add src/ms365_intent_mcp/intent/resolve/impl.py tests/intent/test_resolve_impl.py
git commit -m "feat(#49): thread output_dir through resolve intent impl"
```

---

## Task 8: Regenerate schema snapshots + version bump + full suite

**Files:**
- Modify: `tests/intent/snapshots/schemas/resolve_payload.json`, `resolved_content.json`, `resolved_content_data_union.json`
- Modify: `pyproject.toml`

**Interfaces:** none (release mechanics).

- [ ] **Step 1: Inspect how snapshots are generated**

Run: `uv run pytest tests/intent/test_schema_snapshots.py -v`
Expected: FAIL — the new `attachments`/`output_dir` fields diverge from the committed snapshots. Read `tests/intent/test_schema_snapshots.py` to find the regeneration mechanism (an env flag like `UPDATE_SNAPSHOTS=1`, or a direct dump). Confirm before regenerating.

- [ ] **Step 2: Regenerate the snapshots**

Use the mechanism identified in Step 1 (commonly):

Run: `UPDATE_SNAPSHOTS=1 uv run pytest tests/intent/test_schema_snapshots.py -v`
Then inspect the diff:
Run: `git diff tests/intent/snapshots/schemas/`
Expected: `resolve_payload.json` gains `output_dir`; `resolved_content*.json` gain `attachments` / `AttachmentInfo`. No unrelated churn.

If no auto-update mechanism exists, regenerate by reading the test's dump logic and writing the JSON via the same code path — do NOT hand-edit the JSON.

- [ ] **Step 3: Bump the version**

In `pyproject.toml`, change `version = "1.5.0"` → `version = "1.6.0"`.

- [ ] **Step 4: Run the full suite**

Run: `uv run pytest -q`
Expected: PASS, no failures, no errors.

Then run the linter if configured:
Run: `uv run ruff check src/ tests/` (skip if ruff isn't part of the repo)
Expected: clean, or only pre-existing warnings unrelated to this change.

- [ ] **Step 5: Commit**

```bash
git add tests/intent/snapshots/schemas/ pyproject.toml
git commit -m "chore(#49): regenerate schema snapshots, bump 1.5.0 -> 1.6.0"
```

---

## Task 9: Live smoke test + PR

**Files:** none (verification + integration).

- [ ] **Step 1: Install the branch build**

Per CLAUDE.md deployment rules, test on the branch before any merge:

Run: `uv cache clean ms365-intent-mcp && uv tool install --force git+<repo-url>@feat/49-email-attachments`
Expected: installs `1.6.0`.

- [ ] **Step 2: Manual smoke (requires a real email with inline images)**

Resolve an email URL known to carry inline screenshots, once without and once with `output_dir`. Confirm: metadata lists the images (correct `cid`); with `output_dir` the files land on disk and open as valid images; a plain no-attachment email still resolves with no extra latency.

Report the actual result (files written, sizes) — do not claim success without observing it.

- [ ] **Step 3: Open the PR (review, not auto-merge)**

```bash
git push -u origin feat/49-email-attachments
gh pr create --title "feat(#49): extract inline images + attachments from emails via resolve" \
  --body "Closes #49. Adds attachment enumeration + on-demand download to resolve. Spec: docs/superpowers/specs/2026-07-22-email-attachments-design.md. Reviewed by gpt-4.1 + gemini-2.5-pro; Graph behavior verified against MS Learn. Awaiting live-session confirmation before merge per DEPLOYMENT.md."
```

- [ ] **Step 4: Wait for user confirmation before merge** (DEPLOYMENT.md: never merge Playground repos to main without explicit approval in a live session).

---

## Self-Review

**Spec coverage:**
- Enumerate (gated on `hasAttachments`/`cid:`) → Task 2 + Task 5. ✓
- Metadata-only default, `output_dir` downloads → Tasks 5, 7. ✓
- Inline CID images + regular files (same `fileAttachment` type) → Task 1 classification. ✓
- `$value` fallback for null `contentBytes` → Task 3. ✓
- Defensive pagination → Task 2. ✓
- Size caps (per-attachment + per-request) → Task 3. ✓
- Filename collision suffix, empty/degenerate/255-char, path containment → Tasks 1, 3. ✓
- `itemAttachment`/`referenceAttachment` surfaced with note, never downloaded → Tasks 1, 3. ✓
- Content-type/extension mismatch note → Task 3. ✓
- Graceful degradation on enumeration failure → Tasks 2, 5. ✓
- `AttachmentInfo` model + `EmailContent`/`ResolvePayload` extension → Task 4. ✓
- Formatter section → Task 6. ✓
- `output_dir` at transcript parity, no sandbox → Task 3 (`_resolve_output_base`). ✓
- Version 1.5.0 → 1.6.0 → Task 8. ✓
- Testing matrix → Tasks 1–7 cover every spec test bullet. ✓

**Placeholder scan:** No TBD/TODO. Every code step shows full code. `_human_size` has an explicit "use the plain-loop version" instruction resolving its one ambiguity.

**Type consistency:** metadata dict keys (`name`, `content_type`, `size`, `is_inline`, `cid`, `attachment_id`, `kind`, `_content_bytes`, `note`, `local_path`) are identical across `classify_attachment` (T1), `download_attachments` (T3), and the `_build_structured_data` strip (T5, removing `_content_bytes` + `kind`). `AttachmentInfo` (T4) fields match the stripped dict exactly. `compose_resolve` 4-arg signature consistent across T5 and T7. `enumerate_attachments` returns `(list, str|None)` consistently in T2 and T5.
