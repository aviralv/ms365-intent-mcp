# Email attachment extraction on `resolve` — design

**Issue:** #49 — Cannot extract inline (CID-referenced) images from email bodies
**Date:** 2026-07-22
**Version target:** 1.5.0 → 1.6.0 (minor — new capability, backward-compatible)
**Review:** critiqued by gpt-4.1 and gemini-2.5-pro; Graph-API claims independently
verified against Microsoft Learn (see Graph API mechanics). The models converged on
three issues — `$value` fallback, pagination, `output_dir` write-scope — but their
shared premise that `contentBytes` is omitted above 4 MB on *read* is wrong (that
limit is upload-only). The spec adopts the defensive `$value`/pagination/size-cap
handling on verified grounds, and keeps `output_dir` at parity with the shipped
`transcript` tool rather than the reviewers' multi-tenant sandbox.

## Problem

When an email contains inline images pasted into the body (CID-referenced, e.g.
`[cid:image001.png@01DD19B1.CD8E1710]`), the MCP server exposes the body text and
the *positions* where each image sits — but provides no way to access or download
the image bytes. `find`, `resolve`, and `whats_new` all return the CID placeholder
tokens verbatim; there is no tool surface to fetch the referenced image content.

Inline images are frequently the *substance* of an email — screenshots in a bug
report, diagrams, annotated flows. Today the user must manually drag each image out
of Outlook.

**Confirmed gap (verified across `src/` 2026-07-22):** there is *no* attachment
byte-download handling anywhere in the codebase, for inline images *or* regular
file attachments (PDFs, docx, etc.). The only "attachment" code in `resolve.py`
handles Teams message attachments (forwarded-message text, reply-quote context,
reference-type shared-file *links* — no bytes). Email `hasAttachments` is never
read. The only byte-download path is `transcript` (VTT via `VroomClient`) and
`GraphClient.get_content`. Because the enumeration endpoint
(`/messages/{id}/attachments`) is identical for inline and regular attachments,
this design closes both gaps in one stroke.

## Scope

**In scope:**
- Enumerate email attachments (metadata) on `resolve` of an email URL.
- Download `fileAttachment` bytes to disk on demand (`output_dir`).
- Covers both inline CID images (`isInline=true`) and regular file attachments
  (`isInline=false`).

**Out of scope:**
- `itemAttachment` (embedded email/event/contact — no `contentBytes`, needs
  `$expand`). Surfaced as a metadata row with a note; never downloaded.
- `referenceAttachment` (cloud links — no bytes). Surfaced as a metadata row with
  a note; link resolution is already partially handled elsewhere.
- Non-email URL types — `output_dir` is a no-op for them (unchanged behavior).
- No new tool. Honors the 9-tool ceiling (CLAUDE.md design constraint).

## Graph API mechanics

> Verified against Microsoft Learn (2026-07-22). Sources:
> [fileAttachment resource](https://learn.microsoft.com/en-us/graph/api/resources/fileattachment?view=graph-rest-1.0),
> [message: list attachments](https://learn.microsoft.com/en-us/graph/api/message-list-attachments?view=graph-rest-1.0),
> [attachment: get](https://learn.microsoft.com/en-us/graph/api/attachment-get?view=graph-rest-1.0),
> [attachmentItem](https://learn.microsoft.com/en-us/graph/api/resources/attachmentitem?view=graph-rest-1.0).

`GET /me/messages/{id}/attachments` returns attachments with `@odata.type`:

| Type | `contentBytes`? | `$value` bytes? | Fields used | Handling |
|---|---|---|---|---|
| `#microsoft.graph.fileAttachment` | usually (base64) | yes (raw binary) | `name`, `contentType`, `size`, `isInline`, `contentId` | Enumerate + download |
| `#microsoft.graph.itemAttachment` | no | MIME, not a file | `name`, `size` | Enumerate only (note) |
| `#microsoft.graph.referenceAttachment` | no | 405 (unsupported) | `name` | Enumerate only (note) |

**Verified facts (correcting the reviewers' assumptions):**

- **No documented read-time size cutoff for `contentBytes`.** The 3 MB / 4 MB
  boundary in Graph docs applies *only to the upload/send path*, not to reading.
  Attachments of at least 11 MB read back via `contentBytes` on a GET. So the
  reviewers' "large attachments silently fail because `contentBytes` is omitted
  above 4 MB" is **not accurate for the read path.**
- **`contentBytes` can still be null on a `fileAttachment`** for other reasons: a
  `$select` that excludes it, or a base-type read. Defensive design therefore still
  needs a fallback — but it fires on "null despite `fileAttachment` + non-zero
  `size`", not on a size threshold.
- **`$value` fallback:** `GET /me/messages/{id}/attachments/{attachmentId}/$value`
  returns the raw binary for a `fileAttachment` (Content-Type = the attachment's
  own). Returns MIME for `itemAttachment`, and **405** for `referenceAttachment`.
  Stream this to disk rather than buffering.
- **Pagination is NOT documented** for the attachments collection — no stated page
  size, no guaranteed single-call completeness. Defensive approach: request with a
  high `$top` and follow `@odata.nextLink` if present. Don't hard-code a page-size
  assumption (the reviewers guessed 10 and 100 respectively; docs state neither).
- **`contentId` ↔ `cid:` confirmed.** The HTML body references
  `<img src="cid:<contentId>">`; the inline attachment's `contentId` holds that
  token. `isInline: true` marks inline attachments. Match defensively: strip angle
  brackets and compare case-insensitively (some clients bracket the CID).

## Tool surface

`resolve` gains one optional payload field: `output_dir: str | None = None`.

**Behavior on an email URL:**

1. **Enumeration (always, conditionally):** After the existing message fetch, if
   the email has `hasAttachments=true` **or** the body contains a `cid:` reference,
   fetch `/messages/{id}/attachments` (with a high `$top`, following
   `@odata.nextLink` defensively). Otherwise skip it — emails with no attachments
   cost zero extra Graph calls (the common case stays as fast as today). Populate
   `EmailContent.attachments` with metadata (no bytes).

2. **Download (on demand):** If `output_dir` is passed, materialize each
   `fileAttachment` to disk:
   - If `contentBytes` is present, decode and write.
   - If `contentBytes` is null but it's a `fileAttachment` with non-zero `size`,
     fall back to `GET .../attachments/{id}/$value` and **stream** the raw bytes to
     disk (via `GraphClient.get_content`) rather than buffering.
   - Enforce size caps *before* writing (see Resource limits below).
   Each written entry gains a `local_path`. `itemAttachment`/`referenceAttachment`
   get no `local_path` (a `$value` call on a referenceAttachment returns 405).

**Non-email URL types:** ignore `output_dir` (no-op).

## Data shapes

New `AttachmentInfo` model (in `intent/resolve/schemas.py`):

```
AttachmentInfo:
    name: str
    content_type: str        # "" if unknown
    size: int                # bytes, 0 if unknown
    is_inline: bool
    cid: str                 # contentId for inline images; "" for regular files
    attachment_id: str       # Graph attachment id
    local_path: str | None   # set only after a successful download
    note: str | None         # per-entry note: non-file type, size-capped, $value
                             # failure, content-type/extension mismatch, etc.
```

`EmailContent` gains: `attachments: list[AttachmentInfo] = []`.

`ResolvePayload` gains: `output_dir: str | None = None`.

## Architecture placement

- **New module `composers/attachments.py`** with two focused helpers:
  - `async def enumerate_attachments(client, message_endpoint) -> list[dict]` —
    fetches `/attachments` (high `$top`, follows `@odata.nextLink`), classifies each
    by `@odata.type`, returns metadata dicts (retaining raw `contentBytes`
    internally for the download step when present, but NOT surfacing it in the
    response model).
  - `async def download_attachments(client, message_endpoint, entries, output_dir)`
    — for each `fileAttachment`, decode inline `contentBytes` if present, else
    stream `.../attachments/{id}/$value` to disk via `client.get_content`. Writes
    files with collision-safe names + path-containment guard, enforces size caps,
    sets `local_path`. **Not** pure w.r.t. Graph — the `$value` fallback needs the
    client (this corrects the first-draft assumption that download needed no
    network).
- **`composers/resolve.py`** email branch (`_fetch_resolved` + `_build_structured_data`):
  conditionally enumerate (gated on `hasAttachments`/`cid:`), and if `output_dir`
  was passed, download. Keep the branch thin — logic lives in the new module.
  `output_dir` threads through `compose_resolve(...)`.
- **`resolver.py`** — unchanged. No URL-parsing change; the email endpoint already
  resolves. The `$value` sub-path is built from the resolved message endpoint +
  `/attachments/{id}/$value`, not parsed from a URL.
- **`intent/resolve/schemas.py`** — add `AttachmentInfo`, extend `EmailContent` and
  `ResolvePayload`.
- **`intent/resolve/impl.py`** — thread `payload.output_dir` into `compose_resolve`.
- **`formatters.py`** — email render (`format_resolved_content_markdown`) gains an
  attachments section: 📎 per attachment with name · size · inline/file marker ·
  `local_path` when downloaded, or the `note` for non-file types.

## Resource limits (decode-bomb / oversized guard)

Both reviewers flagged unbounded byte handling. Guards, with defaults as module
constants:

- **Per-attachment cap** (`MAX_ATTACHMENT_BYTES`, default 100 MB): skip an
  attachment whose `size` exceeds it — per-entry `note`, no `local_path`, no fetch.
- **Per-request total cap** (`MAX_TOTAL_ATTACHMENT_BYTES`, default 250 MB): sum
  `size` across qualifying attachments before downloading; if the total exceeds the
  cap, download up to the cap and note the remainder as skipped. Prevents a single
  resolve from writing arbitrarily large volumes.
- The `$value` path buffers the full response in memory (bounded by the 100 MB
  per-file cap); true streaming to disk without a full in-memory buffer is deferred
  as a future improvement.

## Filename collision handling

Multiple attachments in one email can share a filename (inline images are often
`image001.png`, `image002.png`, ...). **De-dupe with suffix:** the first writes
`image001.png`, a second collision writes `image001-2.png`, etc. Nothing is silently
overwritten.

- Sanitize filenames (reuse transcript's `_SANITIZE_RE` approach: `[^A-Za-z0-9._-]+`
  → `_`).
- **Empty / degenerate names:** if the sanitized name is empty, or is `.`/`..`, or
  is all-dots, fall back to `attachment_{n}` (using the 1-based index). Preserve the
  original extension when derivable from `name`/`contentType`.
- **Length:** truncate the final filename to 255 chars (preserving the extension).
- Path-containment guard: resolved destination must stay under `output_dir` (reuse
  transcript's resolved-parent check). Rejects traversal from a hostile filename.
- `output_dir` defaults to `~/.cache/ms365-intent-mcp/attachments` and is created if
  absent (`mkdir(parents=True)`), matching the `transcript` tool's precedent. No
  base-dir sandbox — consistent with `transcript`; this is a local single-user
  server, not a multi-tenant HTTP service. (The reviewers' "arbitrary write"
  concern assumes the latter threat model.)

## Error handling

- **Enumeration failure (403/404/transport):** degrade gracefully. Email body still
  renders; `attachments` is empty; a `_note` explains the enumeration failed. Never
  fail the whole resolve because attachments couldn't be listed.
- **Single-attachment decode/write/`$value` failure:** does not abort the batch.
  That entry gets no `local_path` and a per-entry `note` with the reason; other
  attachments still download.
- **`contentBytes` null on a `fileAttachment` with non-zero `size`:** not an error —
  fall back to `$value` streaming. Only if *that* fails does the entry get a note.
- **`output_dir` points at an existing file (not a dir):** surfaced as an error
  note; no partial writes.
- **Content-type / extension mismatch** (e.g. `contentType` image but `name` ends
  `.pdf`): informational `note` on the entry; file still written under its original
  name (gpt-4.1 suggestion — surfaced, not enforced).

## Testing

Composer unit tests (new `tests/test_composers_attachments.py` for the helpers,
plus `tests/test_composers_resolve.py` for the email-branch wiring):
- CID-only email → inline images enumerated with correct `cid` (incl. angle-bracket
  + case-insensitive CID match).
- Regular-files-only email → enumerated, `cid` empty, `is_inline` false.
- Mixed inline + regular → both enumerated.
- No attachments, no `cid:` in body → **no** extra Graph call fires (assert call
  count).
- `hasAttachments=false` but body has `cid:` → enumeration still fires.
- `itemAttachment` / `referenceAttachment` → surfaced as metadata rows with note,
  never a `local_path`, never crash.
- **`contentBytes` null + non-zero `size` → `$value` fallback fires and streams**
  (assert the `$value` endpoint is hit and bytes land on disk).
- **Pagination:** a two-page `/attachments` response (`@odata.nextLink`) →
  attachments from both pages enumerated.
- **Per-attachment cap:** oversized `size` → skipped with note, no fetch.
- **Per-request total cap:** attachments summing over the cap → remainder skipped
  with note.
- Collision → second file suffixed `-2`.
- Empty/degenerate/over-255-char filename → falls back / truncates correctly.
- Download to tmpdir → correct bytes on disk, `local_path` set.
- Enumeration 403 → body renders, empty attachments, `_note` present.
- Path-traversal filename → contained (writes under `output_dir`, not outside).

Schema tests:
- Snapshot updates for `resolved_content.json`, `resolved_content_data_union.json`,
  `resolve_payload.json` (new fields).

Intent-impl test (`tests/intent/test_resolve_impl.py`):
- `output_dir` round-trip: payload field threads through, `local_path` present in
  response.

## Versioning

`pyproject.toml`: `1.5.0` → `1.6.0`. New backward-compatible capability (new optional
field, new response field defaulting to empty). No breaking change to existing
callers.
