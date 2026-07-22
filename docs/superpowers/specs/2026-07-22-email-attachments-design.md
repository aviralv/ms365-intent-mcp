# Email attachment extraction on `resolve` — design

**Issue:** #49 — Cannot extract inline (CID-referenced) images from email bodies
**Date:** 2026-07-22
**Version target:** 1.5.0 → 1.6.0 (minor — new capability, backward-compatible)

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

`GET /me/messages/{id}/attachments` returns attachments with `@odata.type`:

| Type | Has `contentBytes`? | Fields used | Handling |
|---|---|---|---|
| `#microsoft.graph.fileAttachment` | yes (base64) | `name`, `contentType`, `size`, `isInline`, `contentId` | Enumerate + download |
| `#microsoft.graph.itemAttachment` | no | `name`, `size` | Enumerate only (note) |
| `#microsoft.graph.referenceAttachment` | no | `name` | Enumerate only (note) |

- `contentId` matches the `cid:` token in the body — this is the re-association key
  for inline images.
- `contentBytes` is returned inline with the enumeration call — no separate
  per-attachment download round-trip is needed.

## Tool surface

`resolve` gains one optional payload field: `output_dir: str | None = None`.

**Behavior on an email URL:**

1. **Enumeration (always, conditionally):** After the existing message fetch, if
   the email has `hasAttachments=true` **or** the body contains a `cid:` reference,
   make one extra Graph call to `/messages/{id}/attachments`. Otherwise skip it —
   emails with no attachments cost zero extra Graph calls (the common case stays as
   fast as today). Populate `EmailContent.attachments` with metadata (no bytes).

2. **Download (on demand):** If `output_dir` is passed, decode each
   `fileAttachment`'s `contentBytes` and write to disk. Each written entry gains a
   `local_path`. `itemAttachment`/`referenceAttachment` get no `local_path`.

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
    note: str | None         # e.g. "embedded item — not downloadable" for non-file types
```

`EmailContent` gains: `attachments: list[AttachmentInfo] = []`.

`ResolvePayload` gains: `output_dir: str | None = None`.

## Architecture placement

- **New module `composers/attachments.py`** with two focused helpers:
  - `async def enumerate_attachments(client, message_endpoint) -> list[dict]` —
    calls `/attachments`, classifies each by `@odata.type`, returns metadata dicts
    (including raw `contentBytes` retained internally for the download step, but
    NOT surfaced in the response model).
  - `def download_attachments(entries, output_dir) -> list[dict]` — decodes
    `contentBytes`, writes files with collision-safe names + path-containment
    guard, sets `local_path`. Pure w.r.t. Graph (no network); takes already-fetched
    entries.
- **`composers/resolve.py`** email branch (`_fetch_resolved` + `_build_structured_data`):
  conditionally enumerate (gated on `hasAttachments`/`cid:`), and if `output_dir`
  was passed, download. Keep the branch thin — logic lives in the new module.
  `output_dir` threads through `compose_resolve(...)`.
- **`resolver.py`** — unchanged. No URL-parsing change; the email endpoint already
  resolves.
- **`intent/resolve/schemas.py`** — add `AttachmentInfo`, extend `EmailContent` and
  `ResolvePayload`.
- **`intent/resolve/impl.py`** — thread `payload.output_dir` into `compose_resolve`.
- **`formatters.py`** — email render (`format_resolved_content_markdown`) gains an
  attachments section: 📎 per attachment with name · size · inline/file marker ·
  `local_path` when downloaded, or the `note` for non-file types.

## Filename collision handling

Multiple attachments in one email can share a filename (inline images are often
`image001.png`, `image002.png`, ...). **De-dupe with suffix:** the first writes
`image001.png`, a second collision writes `image001-2.png`, etc. Nothing is silently
overwritten.

- Sanitize filenames (reuse transcript's `_SANITIZE_RE` approach: `[^A-Za-z0-9._-]+`
  → `_`).
- Path-containment guard: resolved destination must stay under `output_dir` (reuse
  transcript's resolved-parent check). Rejects traversal from a hostile filename.
- `output_dir` is created if it doesn't exist (`mkdir(parents=True)`), same as
  transcript.

## Error handling

- **Enumeration failure (403/404/transport):** degrade gracefully. Email body still
  renders; `attachments` is empty; a `_note` explains the enumeration failed. Never
  fail the whole resolve because attachments couldn't be listed.
- **Single-attachment decode/write failure:** does not abort the batch. That entry
  gets no `local_path` and a per-entry `note` with the reason; other attachments
  still download.
- **Missing `contentBytes` on a `fileAttachment`** (unexpected API shape): treated
  as a per-entry failure (note, no `local_path`), not a crash.

## Testing

Composer unit tests (`tests/test_composers_resolve.py` or a new
`tests/test_composers_attachments.py`):
- CID-only email → inline images enumerated with correct `cid`.
- Regular-files-only email → enumerated, `cid` empty, `is_inline` false.
- Mixed inline + regular → both enumerated.
- No attachments, no `cid:` in body → **no** extra Graph call fires (assert call
  count).
- `hasAttachments=false` but body has `cid:` → enumeration still fires.
- `itemAttachment` / `referenceAttachment` → surfaced as metadata rows with note,
  never a `local_path`, never crash.
- Collision → second file suffixed `-2`.
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
