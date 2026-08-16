"""Unit tests for ms365_intent_mcp.transcripts parsing logic.

Ported from ferret-transcripts tests/test_core.py — the pure-function layer
(filename regexes, canonical-item selection, URL parsing, Recording metadata,
chat-event extraction, name/ID matching).

Deliberately NOT ported:
  - The four ``resolve_share_url`` monkeypatch tests — this server reuses its
    own async ``/shares`` resolver instead of porting ferret's sync copy.
  - The ``list_chat_recordings`` integration tests (mocked ``requests``) —
    that walk is now async I/O in the composer; its extraction unit
    (``recording_from_message``) is tested directly here.
"""

from __future__ import annotations

import pytest

from ms365_intent_mcp.transcripts import (
    DEFAULT_SEARCH_BUDGET,
    ONE_ON_ONE_NAME_RE,
    RECORDING_NAME_RE,
    TEAMS_FILENAME_RE,
    Recording,
    _select_canonical_items,
    find_match,
    is_recording,
    parse_recording_url,
    recording_from_message,
    tenant_host_from_upn,
)


def make(name: str, created: str = "2026-05-19T11:00:00Z") -> Recording:
    return Recording(
        name=name,
        item_id="01TEST",
        drive_id="drive",
        size=1,
        created=created,
        personal_site="https://example/personal/x",
    )


# ---------- RECORDING_NAME_RE: discovery filter ----------


def test_filter_accepts_standard_recording():
    assert RECORDING_NAME_RE.search("Sprint Review-20260518_130242-Meeting Recording.mp4")


def test_filter_accepts_recording_with_dashes_in_meeting_name():
    assert RECORDING_NAME_RE.search(
        "Weekly - EA Agents PM Meeting-20260518_130242-Meeting Recording.mp4"
    )


def test_filter_rejects_transcript_sibling():
    assert not RECORDING_NAME_RE.search(
        "Weekly - EA Agents PM Meeting-20260420_130322-Meeting Transcript.mp4"
    )


def test_filter_rejects_arbitrary_files_with_recording_in_name():
    assert not RECORDING_NAME_RE.search("Recording notes.docx")
    assert not RECORDING_NAME_RE.search("My Meeting Recording.mp4")  # no date prefix


def test_filter_rejects_manually_renamed_vtt():
    assert not RECORDING_NAME_RE.search("Weekly - EA Agents PM Meeting (1).vtt")


# ---------- ONE_ON_ONE_NAME_RE + is_recording: 1:1 ad-hoc call carve-out ----------


def test_is_recording_accepts_one_on_one_call():
    assert is_recording("Call with Kulkarni, Bawa-20260410_090059-Meeting Transcript.mp4")


def test_is_recording_accepts_scheduled_meeting():
    assert is_recording("Sprint Review-20260518_130242-Meeting Recording.mp4")


def test_is_recording_rejects_arbitrary_call_with_files():
    assert not is_recording("Call with notes.mp4")
    assert not is_recording("Call with Bob from accounting.mov")


def test_one_on_one_name_re_requires_transcript_role():
    assert not ONE_ON_ONE_NAME_RE.search("Call with X-20260101_120000-Meeting Recording.mp4")


# ---------- _select_canonical_items: pair-aware selection ----------


def _it(name: str, item_id: str = "x") -> dict:
    return {"name": name, "id": item_id}


def test_pair_select_prefers_recording_when_both_siblings_present():
    rec = _it("Sprint Review-20260518_130242-Meeting Recording.mp4", "REC")
    tx = _it("Sprint Review-20260518_130242-Meeting Transcript.mp4", "TX")
    chosen = _select_canonical_items([tx, rec])
    assert len(chosen) == 1
    assert chosen[0]["id"] == "REC"


def test_pair_select_keeps_transcript_when_no_recording_sibling():
    tx = _it(
        "Product Talk with Aviral - Alessia Cesario-20260527_113902-Meeting Transcript.mp4",
        "TX",
    )
    chosen = _select_canonical_items([tx])
    assert len(chosen) == 1
    assert chosen[0]["id"] == "TX"


def test_pair_select_keeps_one_on_one_call_transcript():
    tx = _it(
        "Call with Kulkarni, Bawa-20260410_090059-Meeting Transcript.mp4",
        "TX",
    )
    chosen = _select_canonical_items([tx])
    assert len(chosen) == 1
    assert chosen[0]["id"] == "TX"


def test_pair_select_drops_unparseable_filenames():
    chosen = _select_canonical_items([_it("manually-uploaded.mp4")])
    assert chosen == []


def test_pair_select_handles_multiple_meetings():
    a = _it("Sprint Review-20260518_130242-Meeting Recording.mp4", "A")
    b = _it("Standup-20260519_090000-Meeting Recording.mp4", "B")
    chosen = _select_canonical_items([a, b])
    assert {c["id"] for c in chosen} == {"A", "B"}


# ---------- _select_canonical_items: on_reject callback ----------


def test_pair_select_on_reject_unparseable_filename():
    rejects: list[tuple[str, str]] = []
    chosen = _select_canonical_items(
        [_it("manually-uploaded.mp4"), _it("also-not-teams-shape.txt")],
        on_reject=lambda name, reason: rejects.append((name, reason)),
    )
    assert chosen == []
    assert ("manually-uploaded.mp4", "unparseable-filename") in rejects
    assert ("also-not-teams-shape.txt", "unparseable-filename") in rejects


def test_pair_select_on_reject_paired_out_transcript():
    rec = _it("Sprint Review-20260518_130242-Meeting Recording.mp4", "REC")
    tx = _it("Sprint Review-20260518_130242-Meeting Transcript.mp4", "TX")
    rejects: list[tuple[str, str]] = []
    chosen = _select_canonical_items(
        [tx, rec],
        on_reject=lambda name, reason: rejects.append((name, reason)),
    )
    assert len(chosen) == 1
    assert chosen[0]["id"] == "REC"
    assert (
        "Sprint Review-20260518_130242-Meeting Transcript.mp4",
        "paired-out-transcript",
    ) in rejects


def test_pair_select_on_reject_silent_when_no_callback():
    chosen = _select_canonical_items([_it("manually-uploaded.mp4")])
    assert chosen == []  # no exception on missing callback


def test_pair_select_recurring_meeting_different_dates_independent():
    may21 = _it(
        "Product Talk with Aviral - Alessia Cesario-20260521_090448-Meeting Recording.mp4",
        "MAY21",
    )
    may27 = _it(
        "Product Talk with Aviral - Alessia Cesario-20260527_113902-Meeting Transcript.mp4",
        "MAY27",
    )
    chosen = _select_canonical_items([may21, may27])
    assert {c["id"] for c in chosen} == {"MAY21", "MAY27"}


def test_pair_select_same_day_different_timestamps_independent():
    morning = _it("Standup-20260518_090000-Meeting Recording.mp4", "MORN")
    afternoon = _it("Standup-20260518_140000-Meeting Recording.mp4", "AFT")
    chosen = _select_canonical_items([morning, afternoon])
    assert {c["id"] for c in chosen} == {"MORN", "AFT"}


# ---------- Recording.meeting_name ----------


def test_meeting_name_strips_date_and_role_suffix():
    r = make("Sprint Review-20260518_130242-Meeting Recording.mp4")
    assert r.meeting_name == "Sprint Review"


def test_meeting_name_preserves_dashes_in_meeting_name():
    r = make("Weekly - EA Agents PM Meeting-20260518_130242-Meeting Recording.mp4")
    assert r.meeting_name == "Weekly - EA Agents PM Meeting"


def test_meeting_name_falls_back_to_full_name_when_pattern_missing():
    r = make("not-a-teams-file.mp4")
    assert r.meeting_name == "not-a-teams-file.mp4"


def test_meeting_name_for_one_on_one_call():
    r = make("Call with Kulkarni, Bawa-20260410_090059-Meeting Transcript.mp4")
    assert r.meeting_name == "Call with Kulkarni, Bawa"


# ---------- Recording.meeting_date ----------


def test_meeting_date_uses_filename_when_pattern_matches():
    r = make(
        "Placeholder - PMI LeanIX meeting-20260512_124720-Meeting Recording.mp4",
        created="2026-05-19T11:00:00Z",  # 7 days after the meeting
    )
    assert r.meeting_date == "2026-05-12"


def test_meeting_date_matches_created_when_meeting_was_uploaded_promptly():
    r = make(
        "Sprint Review-20260518_130242-Meeting Recording.mp4",
        created="2026-05-18T11:02:43Z",
    )
    assert r.meeting_date == "2026-05-18"


def test_meeting_date_falls_back_to_created_when_filename_unparseable():
    r = make("manually-uploaded-recording.mp4", created="2026-05-19T11:00:00Z")
    assert r.meeting_date == "2026-05-19"


def test_meeting_date_handles_year_month_day_boundaries():
    r = make("New Year Meeting-20260101_000000-Meeting Recording.mp4", created="")
    assert r.meeting_date == "2026-01-01"
    r = make("End of year-20251231_235959-Meeting Recording.mp4", created="")
    assert r.meeting_date == "2025-12-31"


def test_meeting_date_for_one_on_one_call():
    r = make(
        "Call with Kulkarni, Bawa-20260410_090059-Meeting Transcript.mp4",
        created="2026-04-10T09:36:00Z",
    )
    assert r.meeting_date == "2026-04-10"


def test_meeting_date_for_transcript_sibling_still_parses():
    r = make(
        "Weekly - EA Agents PM Meeting-20260420_130322-Meeting Transcript.mp4",
        created="2026-05-19T00:00:00Z",
    )
    assert r.meeting_date == "2026-04-20"


# ---------- TEAMS_FILENAME_RE: full structural match ----------


def test_teams_filename_re_captures_groups():
    m = TEAMS_FILENAME_RE.match(
        "Weekly - EA Agents PM Meeting-20260518_130242-Meeting Recording.mp4"
    )
    assert m is not None
    assert m.group(1) == "Weekly - EA Agents PM Meeting"
    assert m.group(2) == "20260518_130242"
    assert m.group(3) == "Recording"


def test_teams_filename_re_distinguishes_recording_from_transcript():
    m_rec = TEAMS_FILENAME_RE.match("X-20260518_130242-Meeting Recording.mp4")
    m_tx = TEAMS_FILENAME_RE.match("X-20260518_130242-Meeting Transcript.mp4")
    assert m_rec is not None and m_rec.group(3) == "Recording"
    assert m_tx is not None and m_tx.group(3) == "Transcript"


def test_default_search_budget_is_set():
    assert DEFAULT_SEARCH_BUDGET >= 2000


# ---------- tenant_host_from_upn ----------


def test_tenant_host_from_upn_derives_my_host():
    assert tenant_host_from_upn("aviral.vaid@sap.com") == "sap-my.sharepoint.com"


def test_tenant_host_from_upn_honors_env_override(monkeypatch):
    monkeypatch.setenv("MS365_INTENT_TENANT_HOST", "custom-my.sharepoint.com")
    assert tenant_host_from_upn("anyone@anywhere.com") == "custom-my.sharepoint.com"


def test_tenant_host_from_upn_raises_without_tenant(monkeypatch):
    monkeypatch.delenv("MS365_INTENT_TENANT_HOST", raising=False)
    with pytest.raises(RuntimeError, match="Cannot derive tenant"):
        tenant_host_from_upn("no-at-sign")


# ---------- Recording.to_dict() and JSON contract ----------


def make_full(name="Sprint Review-20260518_130242-Meeting Recording.mp4") -> Recording:
    return Recording(
        name=name,
        item_id="01PMSBCWQF",
        drive_id="b!internal_implementation_detail",
        size=12345678,
        created="2026-05-18T11:02:43Z",
        personal_site="https://sap-my.sharepoint.com/personal/jens_x_sap_com",
        web_url="https://sap-my.sharepoint.com/personal/jens_x_sap_com/Documents/Recordings/Sprint%20Review-...mp4",
    )


def test_to_dict_contains_curated_keys():
    expected_keys = {
        "id",
        "meeting_date",
        "meeting_name",
        "filename",
        "organizer_account",
        "personal_site",
        "web_url",
        "created",
        "size_bytes",
    }
    assert set(make_full().to_dict().keys()) == expected_keys


def test_to_dict_excludes_implementation_details():
    d = make_full().to_dict()
    assert "drive_id" not in d
    assert "name" not in d  # exposed as "filename" instead


def test_to_dict_filename_is_raw_name():
    r = make_full()
    d = r.to_dict()
    assert d["filename"] == r.name
    assert d["meeting_name"] != r.name
    assert d["meeting_name"] == "Sprint Review"


def test_to_dict_size_renamed_to_size_bytes():
    d = make_full().to_dict()
    assert d["size_bytes"] == 12345678
    assert "size" not in d


def test_organizer_account_extracts_slug():
    r = make_full()
    assert r.organizer_account == "jens_x_sap_com"


def test_organizer_account_handles_trailing_slash():
    r = Recording(
        name="x",
        item_id="x",
        drive_id="",
        size=0,
        created="",
        personal_site="https://sap-my.sharepoint.com/personal/abc_sap_com/",
    )
    assert r.organizer_account == "abc_sap_com"


def test_organizer_account_empty_for_non_personal_url():
    r = Recording(
        name="x",
        item_id="x",
        drive_id="",
        size=0,
        created="",
        personal_site="https://leanix.sharepoint.com/sites/team-x",
    )
    assert r.organizer_account == ""


# ---------- parse_recording_url ----------


def test_parse_url_extracts_drive_and_item_from_direct_path():
    url = (
        "https://sap-my.sharepoint.com/personal/aviral_vaid_sap_com"
        "/_api/v2.1/drives/b!QS2r6J/items/013KZCUVL"
        "/media/transcripts/abc/streamContent?is=1"
    )
    p = parse_recording_url(url)
    assert p is not None
    assert p.drive_id == "b!QS2r6J"
    assert p.item_id == "013KZCUVL"
    assert p.site_root == "https://sap-my.sharepoint.com/personal/aviral_vaid_sap_com"
    assert p.filename == ""


def test_parse_url_extracts_filename_from_sharable_v_link():
    url = (
        "https://sap-my.sharepoint.com/:v:/r/personal/aviral_vaid_sap_com"
        "/Documents/Recordings/Call%20with%20Kulkarni,%20Bawa-20260410_090059-Meeting%20Transcript.mp4"
        "?csf=1&web=1&nav=foo"
    )
    p = parse_recording_url(url)
    assert p is not None
    assert p.filename == "Call with Kulkarni, Bawa-20260410_090059-Meeting Transcript.mp4"
    assert p.drive_id == ""
    assert p.item_id == ""
    assert p.site_root == "https://sap-my.sharepoint.com/personal/aviral_vaid_sap_com"


def test_parse_url_extracts_filename_from_onedrive_aspx():
    url = (
        "https://sap-my.sharepoint.com/personal/aviral_vaid_sap_com"
        "/_layouts/15/onedrive.aspx?id=%2Fpersonal%2Faviral_vaid_sap_com"
        "%2FDocuments%2FRecordings%2FFoo-20260101_120000-Meeting%20Recording.mp4"
    )
    p = parse_recording_url(url)
    assert p is not None
    assert p.filename == "Foo-20260101_120000-Meeting Recording.mp4"
    assert p.drive_id == ""
    assert p.item_id == ""


def test_parse_url_supports_team_site_direct_ids():
    url = (
        "https://sap.sharepoint.com/sites/team-nova"
        "/_api/v2.1/drives/b!ABC/items/01XYZ"
        "/media/transcripts/foo/streamContent?is=1"
    )
    p = parse_recording_url(url)
    assert p is not None
    assert p.drive_id == "b!ABC"
    assert p.item_id == "01XYZ"
    assert p.site_root == "https://sap.sharepoint.com/sites/team-nova"


def test_parse_url_returns_none_for_garbage():
    assert parse_recording_url("not a url") is None
    assert parse_recording_url("https://example.com/foo") is None
    assert parse_recording_url("") is None


def test_parse_url_returns_none_for_unrecognized_sharepoint_url():
    url = "https://sap-my.sharepoint.com/personal/aviral_vaid_sap_com/SitePages/Home.aspx"
    assert parse_recording_url(url) is None


def test_parse_url_explicitly_rejects_xplatplugins():
    url = (
        "https://sap-my.sharepoint.com/personal/aviral_vaid_sap_com"
        "/_layouts/15/xplatplugins.aspx?uniqueId=33244275-87f3-4439-88f0-4223df553b0c"
    )
    assert parse_recording_url(url) is None


# ---------- parse_recording_url: Shape 4 (Teams recap link) ----------


def test_parse_url_recap_link_extracts_drive_and_item_from_query():
    url = (
        "https://teams.microsoft.com/l/meetingrecap"
        "?driveId=b%21abc123"
        "&driveItemId=01XYZ"
        "&fileUrl=https%3A%2F%2Fsap-my.sharepoint.com%2Fpersonal%2Faviral_vaid_sap_com"
        "%2FDocuments%2FRecordings%2FFoo.mp4"
        "&threadId=19%3Ameeting_abc%40thread.v2"
    )
    p = parse_recording_url(url)
    assert p is not None
    assert p.drive_id == "b!abc123"
    assert p.item_id == "01XYZ"
    assert p.site_root == "https://sap-my.sharepoint.com/personal/aviral_vaid_sap_com"
    assert p.share_url == ""
    assert p.filename == ""


def test_parse_url_recap_link_works_without_fileUrl():  # noqa: N802
    url = "https://teams.microsoft.com/l/meetingrecap?driveId=b%21abc123&driveItemId=01XYZ"
    p = parse_recording_url(url)
    assert p is not None
    assert p.drive_id == "b!abc123"
    assert p.item_id == "01XYZ"
    assert p.site_root == ""  # no fileUrl means no derivable site_root


def test_parse_url_recap_link_without_ids_returns_none():
    url = "https://teams.microsoft.com/l/meetingrecap?threadId=19%3Ameeting_abc%40thread.v2"
    assert parse_recording_url(url) is None


# ---------- parse_recording_url: Shape 5 (`:v:/p/` short share URL) ----------


def test_parse_url_short_share_v_p_link_returns_share_url_only():
    url = (
        "https://sap-my.sharepoint.com/:v:/p/bawa_kulkarni"
        "/IQAtBSzCB2yDRJ2lwQwcjlasAe2IcdQyM2-7ATfGZI2p8-8"
    )
    p = parse_recording_url(url)
    assert p is not None
    assert p.share_url == url
    assert p.drive_id == ""
    assert p.item_id == ""
    assert p.site_root == ""
    assert p.filename == ""


def test_parse_url_short_share_v_p_strips_query_params():
    url = (
        "https://sap-my.sharepoint.com/:v:/p/bawa_kulkarni"
        "/IQAtBSzCB2yDRJ2lwQwcjlasAe2IcdQyM2-7ATfGZI2p8-8"
        "?nav=eyJyZWZlcnJhbEFwcCI6IlRlYW1zRGVza3RvcCJ9&web=1"
    )
    p = parse_recording_url(url)
    assert p is not None
    assert p.share_url == (
        "https://sap-my.sharepoint.com/:v:/p/bawa_kulkarni"
        "/IQAtBSzCB2yDRJ2lwQwcjlasAe2IcdQyM2-7ATfGZI2p8-8"
    )


def test_parse_url_short_share_takes_precedence_over_personal_segment():
    url = "https://sap-my.sharepoint.com/:v:/p/bawa_kulkarni/IQAtBSzCB2yDRJ2lwQwcjlasAe2"
    p = parse_recording_url(url)
    assert p is not None
    assert p.share_url == url
    assert p.site_root == ""


def test_parse_url_v_r_form_still_works_after_shape_5_added():
    url = (
        "https://sap-my.sharepoint.com/:v:/r/personal/aviral_vaid_sap_com"
        "/Documents/Recordings/Foo-20260101_120000-Meeting%20Recording.mp4"
    )
    p = parse_recording_url(url)
    assert p is not None
    assert p.filename == "Foo-20260101_120000-Meeting Recording.mp4"
    assert p.share_url == ""


# ---------- recording_from_message: chat-event extraction ----------


def _recording_event(
    status: str = "success",
    name: str = "Meeting with Bawa Kulkarni-20260601_091225-Meeting Transcript.mp4",
    url: str = "https://sap-my.sharepoint.com/:v:/p/bawa_kulkarni/IQA123",
    call_id: str = "4f7c8071-be50-425e-a292-fadf7e818f12",
    created: str = "2026-06-01T07:28:29.000Z",
) -> dict:
    return {
        "createdDateTime": created,
        "messageType": "unknownFutureValue",
        "from": None,  # system message has from=None
        "body": {"content": "<systemEventMessage/>"},
        "eventDetail": {
            "@odata.type": "#microsoft.graph.callRecordingEventMessageDetail",
            "callId": call_id,
            "callRecordingDisplayName": name,
            "callRecordingUrl": url,
            "callRecordingStatus": status,
        },
    }


def test_recording_from_message_extracts_success_event():
    msg = _recording_event()
    rec = recording_from_message(msg)
    assert rec is not None
    assert rec.name == "Meeting with Bawa Kulkarni-20260601_091225-Meeting Transcript.mp4"
    assert rec.item_id == "4f7c8071-be50-425e-a292-fadf7e818f12"
    assert rec.web_url == "https://sap-my.sharepoint.com/:v:/p/bawa_kulkarni/IQA123"
    assert rec.created == "2026-06-01T07:28:29.000Z"
    assert rec.drive_id == ""
    assert rec.personal_site == ""


def test_recording_from_message_skips_initial_status():
    rejects: list[tuple[str, str]] = []
    msg = _recording_event(status="initial", url="")
    rec = recording_from_message(msg, on_reject=lambda n, r: rejects.append((n, r)))
    assert rec is None
    assert any(r.startswith("not-success-status") for _, r in rejects)


def test_recording_from_message_skips_chunk_finished_status():
    rejects: list[tuple[str, str]] = []
    msg = _recording_event(status="chunkFinished", url="")
    rec = recording_from_message(msg, on_reject=lambda n, r: rejects.append((n, r)))
    assert rec is None
    assert any(r == "not-success-status:chunkFinished" for _, r in rejects)


def test_recording_from_message_rejects_success_with_empty_url():
    rejects: list[tuple[str, str]] = []
    msg = _recording_event(status="success", url="")
    rec = recording_from_message(msg, on_reject=lambda n, r: rejects.append((n, r)))
    assert rec is None
    assert any(r == "no-recording-url" for _, r in rejects)


def test_recording_from_message_skips_non_recording_event_types():
    msg = _recording_event()
    msg["eventDetail"]["@odata.type"] = "#microsoft.graph.callEndedEventMessageDetail"
    assert recording_from_message(msg) is None


def test_recording_from_message_skips_regular_user_messages():
    msg = {
        "createdDateTime": "2026-06-02T10:00:00Z",
        "messageType": "message",
        "body": {"content": "Hey, did you join the call?"},
        "eventDetail": None,
    }
    assert recording_from_message(msg) is None


def test_recording_from_message_emits_unparseable_filename_warning():
    rejects: list[tuple[str, str]] = []
    msg = _recording_event(name="some-non-standard-name.mp4")
    rec = recording_from_message(msg, on_reject=lambda n, r: rejects.append((n, r)))
    assert rec is not None
    assert rec.name == "some-non-standard-name.mp4"
    assert any(r == "unparseable-filename" for _, r in rejects)


# ---------- find_match (ambiguous prefix detection) ----------


def _make_rec_with_id(item_id: str, name: str = "Meeting") -> Recording:
    return Recording(
        name=f"{name}-20260101_120000-Meeting Recording.mp4",
        item_id=item_id,
        drive_id="d",
        size=1,
        created="2026-01-01T12:00:00Z",
        personal_site="https://example/personal/x",
    )


def test_find_match_exact_id_wins():
    a = _make_rec_with_id("abc")
    b = _make_rec_with_id("abcdef")
    match, ambiguous = find_match([a, b], "abc")
    assert match is a
    assert ambiguous == []


def test_find_match_unique_prefix():
    a = _make_rec_with_id("abcdef123")
    b = _make_rec_with_id("xyz")
    match, ambiguous = find_match([a, b], "abc")
    assert match is a
    assert ambiguous == []


def test_find_match_ambiguous_prefix_returns_candidates():
    a = _make_rec_with_id("abc-123")
    b = _make_rec_with_id("abc-456")
    c = _make_rec_with_id("xyz")
    match, ambiguous = find_match([a, b, c], "abc")
    assert match is None
    assert {r.item_id for r in ambiguous} == {"abc-123", "abc-456"}


def test_find_match_falls_back_to_meeting_name():
    a = _make_rec_with_id("uuid-a", name="Sprint Review")
    b = _make_rec_with_id("uuid-b", name="Backlog Refinement")
    match, ambiguous = find_match([a, b], "Sprint")
    assert match is a
    assert ambiguous == []


def test_find_match_no_match_returns_none_none():
    match, ambiguous = find_match([_make_rec_with_id("xyz")], "nonexistent")
    assert match is None
    assert ambiguous == []


# ---------- find_match (name recency-rank + non-fatal alternatives, #34) ----------


def _make_rec_dated(item_id: str, meeting_name: str, date_yyyymmdd: str) -> Recording:
    """Recording whose Teams filename encodes a specific meeting date."""
    return Recording(
        name=f"{meeting_name}-{date_yyyymmdd}_120000-Meeting Recording.mp4",
        item_id=item_id,
        drive_id="d",
        size=1,
        created=f"{date_yyyymmdd[:4]}-{date_yyyymmdd[4:6]}-{date_yyyymmdd[6:8]}T12:00:00Z",
        personal_site="https://example/personal/x",
    )


def test_find_match_name_multimatch_returns_most_recent():
    """A stale 2-week-old file must never beat a same-day one (the incident)."""
    stale = _make_rec_dated("old", "Meeting with Bawa Kulkarni", "20260701")
    fresh = _make_rec_dated("new", "Meeting with Bawa Kulkarni", "20260714")
    match, alternatives = find_match([stale, fresh], "Bawa")
    assert match is fresh


def test_find_match_name_multimatch_surfaces_alternatives_non_fatally():
    """Multiple name matches return the freshest AS the match (not None) plus
    the losers as alternatives — a soft signal, not a hard ambiguity error."""
    stale = _make_rec_dated("old", "Meeting with Bawa Kulkarni", "20260701")
    fresh = _make_rec_dated("new", "Meeting with Bawa Kulkarni", "20260714")
    match, alternatives = find_match([stale, fresh], "Bawa")
    assert match is fresh
    assert alternatives == [stale]


def test_find_match_name_single_match_has_no_alternatives():
    a = _make_rec_dated("uuid-a", "Sprint Review", "20260714")
    b = _make_rec_dated("uuid-b", "Backlog Refinement", "20260714")
    match, alternatives = find_match([a, b], "Sprint")
    assert match is a
    assert alternatives == []


def test_find_match_prefix_ambiguity_still_hard_error():
    """ID-prefix ambiguity stays a hard error (match None) — distinct from the
    soft name-multimatch signal — so the caller can branch on `match is None`."""
    a = _make_rec_with_id("abc-123")
    b = _make_rec_with_id("abc-456")
    match, candidates = find_match([a, b], "abc")
    assert match is None
    assert {r.item_id for r in candidates} == {"abc-123", "abc-456"}
