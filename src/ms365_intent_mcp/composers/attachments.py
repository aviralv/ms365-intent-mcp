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
from ._utils import _error_reason

ATTACHMENT_FILE_TYPE = "#microsoft.graph.fileAttachment"
_ITEM_TYPE = "#microsoft.graph.itemAttachment"
_REFERENCE_TYPE = "#microsoft.graph.referenceAttachment"

MAX_ATTACHMENT_BYTES = 100 * 1024 * 1024
MAX_TOTAL_ATTACHMENT_BYTES = 250 * 1024 * 1024

_SANITIZE_RE = re.compile(r"[^A-Za-z0-9._-]+")
_MAX_FILENAME_LEN = 255
_CID_RE = re.compile(r"cid:", re.IGNORECASE)
_ENUM_MAX_PAGES = 5


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
    # Collapse consecutive dots (path-traversal guard: ".." → "_")
    cleaned = re.sub(r"\.{2,}", "_", cleaned).strip("_")
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


def body_has_cid(body_text: str) -> bool:
    """True if the body references an inline attachment via a cid: token.

    Works on both HTML bodies (``src="cid:..."``) and the plain-text body Graph
    returns under the ``outlook.body-content-type="text"`` Prefer header, where
    inline refs survive as ``[cid:...]`` tokens.
    """
    return bool(body_text) and _CID_RE.search(body_text) is not None


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

        if not e.get("_content_bytes") and size == 0:
            e["note"] = "skipped — no content (zero-size attachment with no inline bytes)"
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
            reason = _error_reason(exc)
            return entries, f"attachment enumeration failed ({exc.status_code}): {reason}"
        for raw in resp.get("value", []):
            entries.append(classify_attachment(raw))
        next_url = resp.get("@odata.nextLink") or None
        pages += 1
    return entries, None
