"""Pure functions for SharePoint Vroom transcript discovery.

Ported verbatim from ferret-transcripts (`ferret_transcripts/core.py` and
`config.py`) — the string-parsing, canonical-item selection, and filename
regex layer. All I/O lives in ``vroom.py`` (``VroomClient``) and the
``composers/transcript.py`` composer; nothing here touches the network.

Deliberately NOT ported: ferret's ``resolve_share_url`` — this server already
has an async ``/shares`` resolver (``composers.resolve._enrich_call_recording``)
that the composer reuses instead.
"""

from __future__ import annotations

import os
import re
from dataclasses import dataclass
from typing import Callable, Optional
from urllib.parse import parse_qs, unquote, urlparse

# Teams-generated filename pattern, e.g.
#   "Sprint Review-20260518_130242-Meeting Recording.mp4"
# Group 1 captures the meeting name prefix; group 2 the YYYYMMDD_HHMMSS
# timestamp; group 3 the role suffix (Recording for parents, Transcript
# for siblings). The full timestamp (not just the date) is in group 2 so
# back-to-back recordings of the same meeting on the same day can be
# distinguished as independent sibling groups by _select_canonical_items.
# Assumes English-language Teams clients; non-English clients localize the
# "Meeting Recording" string and would not match.
TEAMS_FILENAME_RE = re.compile(
    r"^(.+?)-(\d{8}_\d{6})-Meeting (Recording|Transcript)\."
)
RECORDING_NAME_RE = re.compile(r"-\d{8}_\d{6}-Meeting Recording\.")

# Teams 1:1 ad-hoc call recordings follow the same convention as scheduled
# meetings (`<Name>-YYYYMMDD_HHMMSS-Meeting <Role>.<ext>`), but role is always
# `Transcript` and there is no `Meeting Recording` sibling — Teams does not
# produce one for ad-hoc calls. The leaf filename always begins with
# `Call with `. Anchoring on that prefix keeps the carve-out from ferret
# Issue #1's "drop Transcript siblings" rule scoped to 1:1 calls only —
# scheduled-meeting Transcript siblings remain filtered out.
ONE_ON_ONE_NAME_RE = re.compile(r"^Call with .+-\d{8}_\d{6}-Meeting Transcript\.")

DEFAULT_SEARCH_BUDGET = 5000


def is_recording(name: str) -> bool:
    return bool(RECORDING_NAME_RE.search(name) or ONE_ON_ONE_NAME_RE.search(name))


def tenant_host_from_upn(upn: str) -> str:
    """Derive the SharePoint ``-my`` host from a UPN.

    e.g. ``user@contoso.com`` → ``contoso-my.sharepoint.com``.

    Honors the ``MS365_INTENT_TENANT_HOST`` override (ported from ferret's
    ``FERRET_TENANT_HOST``). Raises when neither an override nor a derivable
    tenant is available.
    """
    env = os.environ.get("MS365_INTENT_TENANT_HOST")
    if env:
        return env
    domain = upn.split("@")[-1] if "@" in upn else ""
    tenant = domain.split(".")[0] if domain else ""
    if not tenant:
        raise RuntimeError(
            "Cannot derive tenant from UPN. Set MS365_INTENT_TENANT_HOST."
        )
    return f"{tenant}-my.sharepoint.com"


def _select_canonical_items(
    items: list[dict],
    on_reject: Optional[Callable[[str, str], None]] = None,
) -> list[dict]:
    """Pair-aware selection: group Teams sibling pairs by (meeting_name,
    timestamp) and prefer the Meeting Recording item; fall back to the
    Meeting Transcript only when no Recording sibling exists in the input.

    The key uses the full YYYYMMDD_HHMMSS timestamp so two recordings of
    the same meeting at different times of day are independent groups.
    Items whose names don't match TEAMS_FILENAME_RE are dropped.

    Use this for *complete* item sets only (Vroom /children folder
    listings). Graph Search returns hit lists that may include the
    Transcript half of a sibling pair without the Recording half — using
    this helper there would re-introduce ferret Issue #1's wrong-content
    bug, where Transcript siblings of recorded scheduled meetings return
    content from a different recording downstream.

    `on_reject(name, reason)` is invoked for each dropped item; reasons:
      - `unparseable-filename`: didn't match TEAMS_FILENAME_RE
      - `paired-out-transcript`: Transcript sibling lost to a Recording
    """
    parsed: dict[tuple[str, str], dict[str, dict]] = {}
    for item in items:
        name = item.get("name", "")
        m = TEAMS_FILENAME_RE.match(name)
        if not m:
            if on_reject:
                on_reject(name, "unparseable-filename")
            continue
        parsed.setdefault((m.group(1), m.group(2)), {})[m.group(3)] = item
    chosen: list[dict] = []
    for siblings in parsed.values():
        rec = siblings.get("Recording")
        tx = siblings.get("Transcript")
        if rec:
            chosen.append(rec)
            if tx and on_reject:
                on_reject(tx.get("name", ""), "paired-out-transcript")
        elif tx:
            chosen.append(tx)
    return chosen


@dataclass
class Recording:
    """A surfaced meeting recording.

    Has two valid shapes depending on discovery path:

    1. **Resolved**: own-drive (Vroom /children) and Graph Search paths
       populate every field including `personal_site` (the SharePoint host
       root). Download works directly via Vroom.

    2. **Share-only**: chat-anchored discovery emits Recording with
       `personal_site=""` and the SharePoint short share URL in `web_url`.
       The Graph `/shares` hop hasn't run yet, so the host isn't known.
       Download auto-promotes to the `web_url` path which does the resolve.

    `requires_share_resolution` is the discriminator. When True, callers
    that depend on `personal_site` must NOT be called — the URL they build
    collapses to schema-less.
    """

    name: str
    item_id: str
    drive_id: str
    size: int
    created: str
    personal_site: str
    web_url: str = ""

    @property
    def requires_share_resolution(self) -> bool:
        """True when this Recording came from chat-anchored discovery and
        still needs a Graph /shares hop before SharePoint Vroom calls can
        target it. Callers must route through the web_url resolve path for
        these recordings, not list_transcripts(self).
        """
        return not self.personal_site

    @property
    def meeting_name(self) -> str:
        match = TEAMS_FILENAME_RE.match(self.name)
        return match.group(1) if match else self.name

    @property
    def meeting_date(self) -> str:
        """Date the meeting actually happened, parsed from the Teams-generated
        filename.

        Falls back to the OneDrive item's createdDateTime when the filename
        doesn't follow Teams' `-YYYYMMDD_HHMMSS-Meeting Recording.` convention.
        Prefer this over `created[:10]` for display: for delayed uploads, the
        item's createdDateTime trails the meeting by days.
        """
        match = TEAMS_FILENAME_RE.match(self.name)
        if not match:
            return self.created[:10]
        d = match.group(2)
        return f"{d[:4]}-{d[4:6]}-{d[6:8]}"

    @property
    def organizer_account(self) -> str:
        """Stable organizer slug (e.g. 'jens_x_sap_com'), parsed from
        personal_site URL.

        Useful for disambiguating same-named meetings organized by different
        people. Returns "" for non-personal sites (group/team SharePoint).
        """
        m = re.search(r"/personal/([^/]+)/?$", self.personal_site)
        return m.group(1) if m else ""

    def to_dict(self) -> dict:
        """Curated JSON-serializable representation.

        `drive_id` is intentionally omitted (implementation detail used only
        to construct Vroom API URLs); `name` is renamed to `filename`.
        """
        return {
            "id": self.item_id,
            "meeting_date": self.meeting_date,
            "meeting_name": self.meeting_name,
            "filename": self.name,
            "organizer_account": self.organizer_account,
            "personal_site": self.personal_site,
            "web_url": self.web_url,
            "created": self.created,
            "size_bytes": self.size,
        }


@dataclass
class Transcript:
    transcript_id: str
    recording: Recording


@dataclass
class ParsedRecordingUrl:
    """Result of parsing a SharePoint URL pointing at a recording.

    Three shapes of populated fields are valid:
      - `drive_id` + `item_id` populated: caller can hit Vroom directly
        (no further resolution needed).
      - `filename` populated: caller resolves via
        `resolve_item_by_filename` against `site_root`.
      - `share_url` populated: caller resolves via Graph
        `/shares/u!{base64url}/driveItem` (needs a Graph token, not just
        SharePoint). `site_root` may be empty in this case — it gets
        derived from the resolved item's webUrl.
    """

    site_root: str
    drive_id: str = ""
    item_id: str = ""
    filename: str = ""
    share_url: str = ""


def parse_recording_url(url: str) -> Optional[ParsedRecordingUrl]:
    """Parse a SharePoint or Teams URL pointing at a meeting recording.

    Tries each shape in order; the first match wins. Returns None when
    nothing matches — caller should treat that as a user-input error.

    Supported shapes:
      1. Direct `/drives/{driveId}/items/{itemId}/...` (Vroom or Graph)
      2. Sharable `/:v:/r/<site>/Documents/Recordings/<filename>?...`
      3. `_layouts/15/onedrive.aspx?id=<encoded-path>`
      4. `teams.microsoft.com/l/meetingrecap?driveId=...&driveItemId=...`
         — Teams "Copy link" output on the recap chiclet.
      5. `/:v:/p/<user>/<share-id>` short share URL — emitted by Teams
         chat recording events (`callRecordingUrl`). The share-id is
         opaque; resolution requires a Graph hop. Parser sets `share_url`
         only; caller is responsible for the resolution step.

    Both personal sites (`/personal/<user>/`) and team-channel sites
    (`/sites/<team>/`) are supported for shapes 1–3.

    NOT supported: `_layouts/15/xplatplugins.aspx?uniqueId=<guid>`. The
    `uniqueId` there is a SharePoint list-item GUID, not a drive-item ID.
    """
    parsed = urlparse(url)
    if not parsed.scheme.startswith("http"):
        return None

    # Shape 4: Teams recap link. Detected by hostname, not site segment —
    # `teams.microsoft.com` URLs don't have `/personal/` or `/sites/` in
    # them, so we check this BEFORE the site_segment guard below would
    # reject the URL. The recap link encodes everything we need in query
    # params; site_root gets derived from the embedded fileUrl.
    if "teams.microsoft.com/l/meetingrecap" in url:
        qs = parse_qs(parsed.query)
        drive_id = qs.get("driveId", [""])[0]
        item_id = qs.get("driveItemId", [""])[0]
        file_url = unquote(qs.get("fileUrl", [""])[0])
        if drive_id and item_id:
            site_root = ""
            if file_url:
                host_match = re.match(r"^(https://[^/]+)", file_url)
                seg_match = re.search(r"/(personal/[^/]+|sites/[^/]+)", file_url)
                if host_match and seg_match:
                    site_root = f"{host_match.group(1)}/{seg_match.group(1)}"
            return ParsedRecordingUrl(
                site_root=site_root,
                drive_id=drive_id,
                item_id=item_id,
            )
        # Recap URL without IDs (rare, older Teams clients?) — fall through
        # so other shapes get a chance at the URL.

    # Shape 5: SharePoint `:v:/p/` short share. The share-id is an opaque
    # token; we can't extract drive/item without a Graph /shares hop.
    # Detected before the host/site_segment guard because the URL DOES
    # carry a /personal/<user>/ segment that would otherwise be parsed as
    # a site_root — but the trailing share-id is NOT a Recordings filename
    # or a drive/item pair, so existing shapes 1-3 would silently miss it.
    short_share_match = re.search(
        r"^https://[^/]+/:v:/p/[^/]+/[A-Za-z0-9_-]+",
        url,
    )
    if short_share_match:
        # Trim query/fragment off the share URL — Graph /shares is sensitive
        # to whitespace and URL-encoding nuances; the canonical share form
        # ends right after the share-id. Including ?web=1 etc. has been
        # observed to flip the response.
        share_clean = short_share_match.group(0)
        return ParsedRecordingUrl(
            site_root="",  # populated post-resolution from item.webUrl
            share_url=share_clean,
        )

    # Canonical site root: `https://<host>/personal/<user>` or
    # `https://<host>/sites/<team>`. Extract host and site segment
    # separately so we handle both bare URLs and sharable links like
    # `https://<host>/:v:/r/personal/<user>/...` where the :v:/r/
    # prefix sits between the host and the site segment.
    host_match = re.match(r"^(https://[^/]+)", url)
    site_segment_match = re.search(r"/(personal/[^/]+|sites/[^/]+)", url)
    if not host_match or not site_segment_match:
        return None
    site_root = f"{host_match.group(1)}/{site_segment_match.group(1)}"

    pair_match = re.search(r"/drives/([^/]+)/items/([^/?]+)", url)
    if pair_match:
        return ParsedRecordingUrl(
            site_root=site_root,
            drive_id=pair_match.group(1),
            item_id=pair_match.group(2),
        )

    sharable_match = re.search(
        r"/:v:/r/(?:personal/[^/]+|sites/[^/]+)/Documents/Recordings/([^?]+)",
        url,
    )
    if sharable_match:
        return ParsedRecordingUrl(
            site_root=site_root,
            filename=unquote(sharable_match.group(1)),
        )

    if "/_layouts/15/onedrive.aspx" in url:
        qs = parse_qs(parsed.query)
        path = unquote(qs.get("id", [""])[0])
        rec_match = re.search(r"/Recordings/([^/]+)$", path)
        if rec_match:
            return ParsedRecordingUrl(
                site_root=site_root,
                filename=rec_match.group(1),
            )

    return None


def _personal_site_from_web_url(web_url: str) -> str:
    match = re.match(r"(https://[^/]+/personal/[^/]+)", web_url)
    return match.group(1) if match else ""


# Chat-anchored discovery: the discriminator is `eventDetail.@odata.type`
# == `#microsoft.graph.callRecordingEventMessageDetail` with
# `callRecordingStatus == 'success'`. The message-to-Recording extraction is
# pure (dict in, Recording|None out); the chat/message walk that feeds it is
# async I/O and lives in the composer.
CALL_RECORDING_EVENT_TYPE = "#microsoft.graph.callRecordingEventMessageDetail"


def recording_from_message(
    msg: dict,
    on_reject: Optional[Callable[[str, str], None]] = None,
) -> Optional[Recording]:
    """Extract a Recording from a chat message if it's a successful
    recording event, else None.

    The eventDetail shape:
        {
          "@odata.type": "#microsoft.graph.callRecordingEventMessageDetail",
          "callId": "<uuid>",
          "callRecordingDisplayName": "<filename>.mp4",
          "callRecordingUrl": "https://<host>/:v:/p/<user>/<id>" | None,
          "callRecordingStatus": "initial" | "chunkFinished" | "success",
          ...
        }

    Recording fields require adaptation — the chat-event source exposes a
    different shape than Vroom /children or Graph Search:
      - item_id: use callId as a synthetic ID (stable, dedup-friendly). The
        real SharePoint item-id only exists after a /shares hop.
      - drive_id / personal_site: empty. Only resolved post-/shares.
      - web_url: use callRecordingUrl so callers can resolve/download it.
    """
    event_detail = msg.get("eventDetail") or {}
    if event_detail.get("@odata.type", "") != CALL_RECORDING_EVENT_TYPE:
        return None

    name = event_detail.get("callRecordingDisplayName", "") or ""
    status = event_detail.get("callRecordingStatus", "") or ""
    rec_url = event_detail.get("callRecordingUrl", "") or ""
    call_id = event_detail.get("callId", "") or ""

    if status != "success":
        if on_reject:
            on_reject(name or "<unnamed>", f"not-success-status:{status}")
        return None
    if not rec_url:
        if on_reject:
            on_reject(name or "<unnamed>", "no-recording-url")
        return None

    # The displayName follows the standard Teams filename convention; if it
    # doesn't, downstream meeting_name / meeting_date gracefully fall back to
    # the raw name / created timestamp. Surface it as a rejection for debug,
    # but still emit the Recording — the share URL is valid regardless.
    if not TEAMS_FILENAME_RE.match(name) and on_reject:
        on_reject(name or "<unnamed>", "unparseable-filename")

    return Recording(
        name=name,
        item_id=call_id or rec_url,  # callId preferred (stable, dedup-friendly)
        drive_id="",
        size=0,  # not reported in chat-event shape
        created=msg.get("createdDateTime", ""),
        personal_site="",
        web_url=rec_url,
    )


def find_match(
    recordings: list[Recording], target: str
) -> tuple[Optional[Recording], list[Recording]]:
    """Locate a recording by exact ID, ID prefix, or meeting-name substring.

    Returns `(match, ambiguous)`:
      - `(match, [])` — exactly one ID/prefix match OR exactly one
        name-substring match. Caller proceeds.
      - `(None, candidates)` — multiple ID-prefix matches. Caller surfaces
        an ambiguity error so the user can disambiguate with a longer prefix.
      - `(None, [])` — no match at all.

    ID-exact match wins outright. Name-substring matching is fall-back-only
    and keeps first-match-wins semantics — ambiguity detection there would
    fire on every recurring meeting and break ergonomics.
    """
    target_lower = target.lower()
    exact = [r for r in recordings if r.item_id == target]
    if len(exact) == 1:
        return exact[0], []

    prefix = [r for r in recordings if r.item_id.startswith(target)]
    if len(prefix) == 1:
        return prefix[0], []
    if len(prefix) > 1:
        return None, prefix

    for r in recordings:
        if target_lower in r.meeting_name.lower():
            return r, []

    return None, []
