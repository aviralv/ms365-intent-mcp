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
