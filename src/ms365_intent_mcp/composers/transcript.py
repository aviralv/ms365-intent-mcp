"""transcript composer — discover a meeting recording and download its VTT.

Folds ferret-transcripts into the server (issue #29). Three resolution paths:

  * **url** — a recording URL. Fast path when it carries drive/item IDs
    (e.g. ``meeting()``'s ``vroom_url``); short ``:v:/p/`` share URLs take one
    Graph ``/shares`` hop first; sharable/aspx URLs resolve the filename via
    Vroom.
  * **name** — filename discovery across own-drive ``/children`` + Graph Search
    + Teams chat-recording events, then best-match by ID/name. Additive over
    ``meeting()``'s chat-event discovery: catches team-site recordings and
    recordings on other people's drives.

The VTT is streamed to disk; the response carries the path plus meeting
name/date and line/speaker metadata, keeping model context lean.

Reuse note (issue #29): the Graph ``/shares`` share-encoding is reused from
``composers.resolve`` (``_encode_share_url``) rather than porting ferret's sync
``resolve_share_url``. The download/list I/O lives in ``VroomClient``; the
parsing layer in ``transcripts``.
"""

from __future__ import annotations

import asyncio
import re
from pathlib import Path

from ..graph import GraphAPIError, GraphClient
from ..permissions import PermissionRegistry
from ..transcripts import (
    Recording,
    _select_canonical_items,
    find_match,
    is_recording,
    parse_recording_url,
    recording_from_message,
    tenant_host_from_upn,
)
from ..vroom import VroomClient, VroomError
from .resolve import _encode_share_url

GRAPH_SEARCH_BUDGET = 1000  # up to ~5 pages of 200 hits for a name lookup
CHATS_PAGE_SIZE = 50
MAX_CHATS = 200
MESSAGES_PER_CHAT = 30
CHAT_FETCH_CONCURRENCY = 8  # cap parallel per-chat message fetches


class TranscriptResult:
    """Lightweight carrier for the composer's structured output."""

    def __init__(
        self,
        *,
        status: str,
        file_path: str = "",
        meeting_name: str = "",
        meeting_date: str = "",
        line_count: int = 0,
        byte_count: int = 0,
        has_speaker_tags: bool = False,
        message: str = "",
        alternatives_count: int = 0,
    ):
        self.status = status
        self.file_path = file_path
        self.meeting_name = meeting_name
        self.meeting_date = meeting_date
        self.line_count = line_count
        self.byte_count = byte_count
        self.has_speaker_tags = has_speaker_tags
        self.message = message
        self.alternatives_count = alternatives_count


async def compose_transcript(
    graph: GraphClient,
    vroom: VroomClient,
    permissions: PermissionRegistry,
    *,
    url: str | None,
    name: str | None,
    output_dir: str | None,
    item_id: str | None = None,
    drive_id: str | None = None,
    site_root: str | None = None,
    list_recordings: bool = False,
) -> tuple[dict, str]:
    """Resolve a recording and download its VTT. Returns ``(data, markdown)``.

    Five input modes (validated mutually-exclusive upstream in the schema):
      * ``list_recordings=True`` — enumerate discovered recordings, newest
        first; no download. Ferret ``list`` parity, and the escape hatch for
        the counterpart-naming gap (issue #34): an ad-hoc call titled from the
        other side ("Call with Vaid, Aviral") shows up here by date even when
        a counterpart-name search can't match it.
      * ``item_id`` + ``drive_id`` + ``site_root`` — deterministic by-coords
        download, zero discovery (issue #33).
      * ``url`` — a recording URL.
      * ``name`` — filename discovery + best-match.
    """
    scope_msg = permissions.check("Sites.Read.All")
    if scope_msg:
        return _fail(scope_msg)

    if list_recordings:
        return await _compose_list(graph, vroom)

    alternatives: list[Recording] = []
    try:
        if item_id and drive_id and site_root:
            resolved_site, resolved_drive, resolved_item, hint = (
                site_root, drive_id, item_id, "",
            )
        elif url:
            resolved_site, resolved_drive, resolved_item, hint = await _resolve_from_url(
                graph, vroom, url
            )
        elif name:
            resolved_site, resolved_drive, resolved_item, hint, alternatives = (
                await _resolve_from_name(graph, vroom, name)
            )
        else:
            return _fail("Provide `url`, `name`, an item_id+drive_id+site_root triple, or list=true.")
    except VroomError as exc:
        return _fail(_vroom_reason(exc))
    except GraphAPIError as exc:
        return _fail(f"Graph error: {exc.error_code}")

    site_root, drive_id, item_id = resolved_site, resolved_drive, resolved_item

    if not (site_root and drive_id and item_id):
        return _fail(_unresolved_reason(hint, site_root, drive_id, item_id))

    try:
        transcripts = await vroom.list_transcripts(site_root, drive_id, item_id)
    except VroomError as exc:
        return _fail(_vroom_reason(exc))
    if not transcripts:
        return _fail(f"No transcript media found for {hint or 'this recording'}.")

    try:
        dest = _dest_path(output_dir, hint, transcripts[0]["id"])
    except ValueError as exc:
        return _fail(str(exc))
    try:
        byte_count = await vroom.download_transcript_to_file(
            site_root, drive_id, item_id, transcripts[0]["id"], str(dest)
        )
    except VroomError as exc:
        return _fail(_vroom_reason(exc))

    meeting_name, meeting_date = _split_hint(hint)
    line_count, has_tags = _vtt_metadata(dest)

    result = TranscriptResult(
        status="ok",
        file_path=str(dest),
        meeting_name=meeting_name,
        meeting_date=meeting_date,
        line_count=line_count,
        byte_count=byte_count,
        has_speaker_tags=has_tags,
        alternatives_count=len(alternatives),
    )
    return _render(result, alternatives=alternatives)


# ---------------------------------------------------------------------------
# URL path
# ---------------------------------------------------------------------------


async def _resolve_from_url(
    graph: GraphClient, vroom: VroomClient, url: str
) -> tuple[str, str, str, str]:
    """Resolve a recording URL to ``(site_root, drive_id, item_id, hint)``.

    ``hint`` is a display label ("<meeting_name>|<meeting_date>" when derivable,
    else a filename or empty). Raises on transport errors; returns empties in
    the coord slots when the URL can't be resolved (caller surfaces the hint).
    """
    parsed = parse_recording_url(url)
    if parsed is None:
        return "", "", "", f"Could not parse recording URL: {url[:80]}"

    # Short share URL — one Graph /shares hop to get real drive/item + host.
    if parsed.share_url:
        return await _resolve_share(graph, parsed.share_url)

    site_root = parsed.site_root
    drive_id, item_id = parsed.drive_id, parsed.item_id

    if drive_id and item_id:
        if not site_root:
            return "", "", "", (
                "This URL carries drive/item IDs but no site host "
                "(older Teams recap link). Provide the OneDrive 'Copy link' "
                "URL or meeting()'s vroom_url instead."
            )
        return site_root, drive_id, item_id, _filename_hint(url)

    # Filename-only (sharable /:v:/r/ or onedrive.aspx) — resolve via Vroom
    # path-addressing against the site's Recordings folder.
    if parsed.filename:
        drive_id, item_id = await vroom.resolve_item_by_filename(
            parsed.site_root, parsed.filename
        )
        if not (drive_id and item_id):
            return "", "", "", (
                f"'{parsed.filename}' not found in the Recordings folder at "
                f"{parsed.site_root}."
            )
        return parsed.site_root, drive_id, item_id, parsed.filename

    return "", "", "", f"URL didn't contain enough info to locate the recording: {url[:80]}"


async def _resolve_share(graph: GraphClient, share_url: str) -> tuple[str, str, str, str]:
    """Resolve a ``:v:/p/`` short share URL via Graph ``/shares/{ref}/driveItem``.

    Reuses ``_encode_share_url`` from ``composers.resolve``. Derives ``site_root``
    from the resolved item's ``webUrl`` (works for both ``/personal/`` and
    ``/sites/`` recordings). A 403 here is a cross-organizer permission gap.
    """
    ref = _encode_share_url(share_url)
    item = await graph.get(f"/shares/{ref}/driveItem")
    drive_id = (item.get("parentReference") or {}).get("driveId", "")
    item_id = item.get("id", "")
    web_url = item.get("webUrl", "")
    name = item.get("name", "")
    site_root = _site_root_from_web_url(web_url)
    if not (drive_id and item_id and site_root):
        return "", "", "", "Share URL did not resolve to a downloadable recording."
    return site_root, drive_id, item_id, name


# ---------------------------------------------------------------------------
# Name path — 3-source discovery
# ---------------------------------------------------------------------------


async def _discover_all_recordings(
    graph: GraphClient, vroom: VroomClient
) -> list[Recording]:
    """Run 3-source discovery (own-drive + Graph Search + Teams chats), dedupe
    by item_id, and return newest-first.

    Shared by the ``name`` best-match path and the ``list`` enumeration path.
    Own-drive/Search resolved shapes come before chat share-only shapes, so a
    resolved Recording wins dedup over its share-only twin.
    """
    recordings: list[Recording] = []
    recordings.extend(await _discover_own_drive(graph, vroom))
    recordings.extend(await _discover_search(graph))
    recordings.extend(await _discover_chats(graph))

    seen: set[str] = set()
    deduped: list[Recording] = []
    for r in recordings:
        if r.item_id in seen:
            continue
        seen.add(r.item_id)
        deduped.append(r)

    deduped.sort(key=lambda r: r.meeting_date, reverse=True)
    return deduped


async def _compose_list(graph: GraphClient, vroom: VroomClient) -> tuple[dict, str]:
    """List discovered recordings, newest first — ferret ``list`` parity."""
    recordings = await _discover_all_recordings(graph, vroom)
    rows = [
        {
            "meeting_date": r.meeting_date,
            "meeting_name": r.meeting_name,
            "id": r.item_id,
        }
        for r in recordings
    ]
    data = {"status": "ok", "recordings": rows, "message": ""}
    if not rows:
        return data, "No recordings found."
    lines = ["📼 **Recordings** (newest first)", ""]
    lines.append("| Date | Meeting | ID |")
    lines.append("|---|---|---|")
    for row in rows:
        lines.append(f"| {row['meeting_date']} | {row['meeting_name']} | `{row['id']}` |")
    lines.append("")
    lines.append(
        f"{len(rows)} recording(s). Download one with "
        "`transcript(payload={\"name\": \"<meeting>\"})` or its `id`."
    )
    return data, "\n".join(lines)


async def _resolve_from_name(
    graph: GraphClient, vroom: VroomClient, name: str
) -> tuple[str, str, str, str, list[Recording]]:
    """Discover recordings by filename across three sources, best-match ``name``,
    then resolve the match to downloadable coords.

    Returns ``(site_root, drive_id, item_id, hint, alternatives)``. When several
    recordings match the name, ``alternatives`` holds the runner-up matches
    (freshest already chosen as the match) so the caller can surface them — a
    stale pick is never silent (issue #34)."""
    deduped = await _discover_all_recordings(graph, vroom)

    match, alternatives = find_match(deduped, name)
    # A None match with candidates is the hard ID-prefix ambiguity case.
    if match is None and alternatives:
        listing = "; ".join(f"{r.meeting_name} ({r.meeting_date})" for r in alternatives[:5])
        return "", "", "", f"'{name}' is ambiguous — matches: {listing}. Be more specific.", []
    if match is None:
        return "", "", "", f"No recording found matching '{name}'.", []

    hint = f"{match.meeting_name}|{match.meeting_date}"

    # Share-only (chat-discovered) — resolve via the /shares hop.
    if match.requires_share_resolution:
        if not match.web_url:
            return "", "", "", f"Recording '{match.meeting_name}' is missing a resolvable URL.", []
        site_root, drive_id, item_id, note = await _resolve_share(graph, match.web_url)
        # If the share hop failed, surface its note (why) rather than letting
        # the caller collapse to a generic name-echo — issue #31 ask #1.
        if site_root and drive_id and item_id:
            return site_root, drive_id, item_id, hint, alternatives
        return "", "", "", note, []

    return match.personal_site, match.drive_id, match.item_id, hint, alternatives


async def _discover_own_drive(graph: GraphClient, vroom: VroomClient) -> list[Recording]:
    """List the user's own Recordings folder via Vroom /children."""
    try:
        me = await graph.get("/me", params={"$select": "userPrincipalName"})
    except GraphAPIError:
        return []
    upn = me.get("userPrincipalName", "")
    if not upn:
        return []
    try:
        host = tenant_host_from_upn(upn)
    except RuntimeError:
        return []
    username = upn.replace("@", "_").replace(".", "_")
    site_root = f"https://{host}/personal/{username}"

    try:
        items, drive_id = await vroom.list_recordings_children(site_root)
    except VroomError:
        return []

    chosen = _select_canonical_items(items)
    results = [
        Recording(
            name=item.get("name", ""),
            item_id=item["id"],
            drive_id=drive_id,
            size=item.get("size", 0),
            created=item.get("createdDateTime", ""),
            personal_site=site_root,
            web_url=item.get("webUrl", ""),
        )
        for item in chosen
    ]
    results.sort(key=lambda r: r.meeting_date, reverse=True)
    return results


async def _discover_search(graph: GraphClient) -> list[Recording]:
    """Discover recordings via Microsoft Search (driveItem entity).

    Catches recordings on other people's personal drives (meetings organized
    by others where the user participated). Filters to recording parents only.
    """
    body_base = {
        "requests": [
            {
                "entityTypes": ["driveItem"],
                "query": {
                    "queryString": 'filename:"Meeting Recording" OR filename:"Call with"'
                },
                "from": 0,
                "size": 200,
                "sortProperties": [{"name": "createdDateTime", "isDescending": True}],
                "fields": ["id", "name", "createdDateTime", "size", "webUrl", "parentReference"],
            }
        ]
    }

    hits: list[dict] = []
    frm = 0
    while len(hits) < GRAPH_SEARCH_BUDGET:
        body = {"requests": [{**body_base["requests"][0], "from": frm}]}
        try:
            resp = await graph.post("/search/query", json_data=body)
        except GraphAPIError:
            break
        containers = (resp.get("value") or [{}])[0].get("hitsContainers", [{}])
        page_hits = containers[0].get("hits", [])
        hits.extend(page_hits)
        more = containers[0].get("moreResultsAvailable", False)
        if not more or len(page_hits) < 200:
            break
        frm += 200

    results: list[Recording] = []
    for h in hits:
        res = h.get("resource", {})
        name = res.get("name", "")
        if not is_recording(name):
            continue
        drive_id = (res.get("parentReference") or {}).get("driveId", "")
        item_id = res.get("id", "")
        web_url = res.get("webUrl", "")
        site_root = _site_root_from_web_url(web_url)
        if not (drive_id and item_id and site_root):
            continue
        results.append(
            Recording(
                name=name,
                item_id=item_id,
                drive_id=drive_id,
                size=res.get("size", 0),
                created=res.get("createdDateTime", ""),
                personal_site=site_root,
                web_url=web_url,
            )
        )
    return results


async def _discover_chats(graph: GraphClient) -> list[Recording]:
    """Discover recordings via Teams chat callRecording events (share-only).

    No indexing latency — the event posts the moment Teams finishes processing.
    Returns share-only Recordings (personal_site empty; web_url is the share).
    """
    try:
        chats, _ = await graph.get_all(
            f"/me/chats?$top={CHATS_PAGE_SIZE}", max_pages=MAX_CHATS // CHATS_PAGE_SIZE
        )
    except GraphAPIError:
        return []

    chat_ids = [c.get("id", "") for c in chats if c.get("id")]
    if not chat_ids:
        return []

    # Fetch per-chat messages concurrently (bounded) instead of serially —
    # an active user can have 100+ chats, and one-at-a-time is the dominant
    # latency of name-based discovery. Semaphore caps in-flight requests so we
    # don't hammer Graph into 429s.
    sem = asyncio.Semaphore(CHAT_FETCH_CONCURRENCY)

    async def _fetch(chat_id: str) -> list[dict]:
        async with sem:
            try:
                msgs = await graph.get(
                    f"/chats/{chat_id}/messages", params={"$top": str(MESSAGES_PER_CHAT)}
                )
            except GraphAPIError:
                return []
        return msgs.get("value", [])

    per_chat = await asyncio.gather(*(_fetch(cid) for cid in chat_ids))

    # Dedup after gather, iterating in chat order for deterministic results.
    results: list[Recording] = []
    seen_call_ids: set[str] = set()
    for msgs in per_chat:
        for msg in msgs:
            rec = recording_from_message(msg)
            if rec is None or rec.item_id in seen_call_ids:
                continue
            seen_call_ids.add(rec.item_id)
            results.append(rec)
    return results


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _site_root_from_web_url(web_url: str) -> str:
    """Derive ``https://<host>/personal/<user>`` or ``.../sites/<team>`` from a
    driveItem webUrl. Empty when neither segment is present."""
    if not web_url:
        return ""
    host_match = re.match(r"^(https://[^/]+)", web_url)
    seg_match = re.search(r"/(personal/[^/]+|sites/[^/]+)", web_url)
    if host_match and seg_match:
        return f"{host_match.group(1)}/{seg_match.group(1)}"
    return ""


def _filename_hint(url: str) -> str:
    """Best-effort meeting label from a URL's Recordings filename."""
    m = re.search(r"/Recordings/([^/?]+)", url)
    return m.group(1) if m else ""


_SANITIZE_RE = re.compile(r"[^A-Za-z0-9._-]+")


def _dest_path(output_dir: str | None, hint: str, transcript_id: str) -> Path:
    base = (
        Path(output_dir).expanduser()
        if output_dir
        else Path.home() / ".cache" / "ms365-intent-mcp" / "transcripts"
    ).resolve()
    base.mkdir(parents=True, exist_ok=True)
    meeting_name, meeting_date = _split_hint(hint)
    stem = _SANITIZE_RE.sub("_", meeting_name or "transcript").strip("_") or "transcript"
    date_part = f"-{meeting_date}" if meeting_date else ""
    # Sanitize the id fragment too — it feeds the filename, and a stray '/' or
    # '..' from an unexpected API shape must not let the name escape `base`.
    id_frag = _SANITIZE_RE.sub("_", transcript_id[:8]) or "0"
    dest = (base / f"{stem}{date_part}-{id_frag}.vtt").resolve()
    # Containment guard: the resolved filename must stay under `base`. Both are
    # already resolved, so this catches any traversal the sanitizers missed.
    if base != dest.parent:
        raise ValueError("Refusing to write transcript outside the output directory.")
    return dest


def _unresolved_reason(
    hint: str, site_root: str, drive_id: str, item_id: str
) -> str:
    """Build a diagnostic message naming *which* coordinate failed to resolve.

    The old guard echoed the meeting name (``hint``) verbatim, which made a
    resolvable-but-incompletely-coordinated recording look like a discovery
    miss — the opaque error in issue #31. Naming the missing piece
    (site host / drive / item) points at the actual failing hop.
    """
    if not (site_root or drive_id or item_id):
        # Nothing resolved at all — a genuine discovery miss.
        return hint or "Could not locate the recording."
    missing = []
    if not site_root:
        missing.append("site host")
    if not drive_id:
        missing.append("drive id")
    if not item_id:
        missing.append("item id")
    label = hint.split("|", 1)[0] if hint else "this recording"
    return (
        f"Located '{label}' but could not resolve its {', '.join(missing)} — "
        f"the recording was found but is not downloadable via SharePoint."
    )


def _split_hint(hint: str) -> tuple[str, str]:
    """Split a "<meeting_name>|<meeting_date>" hint. Falls back to (hint, "")."""
    if hint and "|" in hint:
        name, date = hint.split("|", 1)
        return name, date
    return hint, ""


def _vtt_metadata(path: Path) -> tuple[int, bool]:
    """Return (line_count, has_speaker_tags) by reading the written VTT."""
    try:
        text = path.read_text(encoding="utf-8", errors="replace")
    except OSError:
        return 0, False
    return text.count("\n") + 1, "<v " in text


def _vroom_reason(exc: VroomError) -> str:
    if exc.status_code == 403:
        return (
            "Access denied (403) — you may need to open this recording in Teams "
            "once before it can be downloaded, or you lack permission to it."
        )
    if exc.status_code == 401:
        return "SharePoint authentication failed (401). Re-run: ms365-intent-mcp auth"
    return f"SharePoint error ({exc.status_code}): {exc.message}"


def _fail(message: str) -> tuple[dict, str]:
    return _render(TranscriptResult(status="error", message=message))


def _render(
    result: TranscriptResult, alternatives: list[Recording] | None = None
) -> tuple[dict, str]:
    data = {
        "status": result.status,
        "file_path": result.file_path,
        "meeting_name": result.meeting_name,
        "meeting_date": result.meeting_date,
        "line_count": result.line_count,
        "byte_count": result.byte_count,
        "has_speaker_tags": result.has_speaker_tags,
        "message": result.message,
        "alternatives_count": result.alternatives_count,
    }
    if result.status != "ok":
        return data, f"❌ {result.message}"

    header = result.meeting_name or "Transcript"
    if result.meeting_date:
        header += f" — {result.meeting_date}"
    tags = "✅ speaker-tagged" if result.has_speaker_tags else "no speaker tags"
    markdown = (
        f"📄 **{header}**\n\n"
        f"- Saved to: `{result.file_path}`\n"
        f"- {result.line_count} lines, {result.byte_count:,} bytes ({tags})"
    )
    # A stale-pick guard (#34): when the name matched more than one recording,
    # name the freshest-first alternatives so the caller can sanity-check that
    # the downloaded one is the intended meeting.
    if alternatives:
        listing = "; ".join(f"{r.meeting_name} ({r.meeting_date})" for r in alternatives[:5])
        markdown += (
            f"\n\n⚠️ {len(alternatives)} other recording(s) also matched — "
            f"downloaded the most recent. Others: {listing}. "
            f"Pass the `id` from `list=true` to pick a specific one."
        )
    return data, markdown
