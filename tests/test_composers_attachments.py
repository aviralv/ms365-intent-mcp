import base64 as _b64
from pathlib import Path

import pytest
from unittest.mock import AsyncMock

from ms365_intent_mcp.composers.attachments import (
    body_has_cid,
    classify_attachment,
    download_attachments,
    enumerate_attachments,
    safe_filename,
    MAX_ATTACHMENT_BYTES,
    _ENUM_MAX_PAGES,
)
from ms365_intent_mcp.graph import GraphAPIError


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
        client.get.assert_awaited()

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

    @pytest.mark.asyncio
    async def test_enum_max_pages_cap(self):
        """Pagination stops at _ENUM_MAX_PAGES even when nextLink is always present."""
        client = AsyncMock()
        infinite_page = {
            "value": [{"@odata.type": "#microsoft.graph.fileAttachment",
                       "name": "a.png", "size": 1, "id": "i1", "contentBytes": "AA=="}],
            "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/messages/M1/attachments?$skip=100",
        }
        client.get = AsyncMock(return_value=infinite_page)
        entries, err = await enumerate_attachments(client, "/me/messages/M1")
        assert client.get.await_count == _ENUM_MAX_PAGES
        assert err is None


def _file_entry(name, cid="", content_bytes: str | None = "AA==", size=1, aid="i", ct="image/png"):
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

    @pytest.mark.asyncio
    async def test_zero_size_no_inline_bytes_skipped_no_value_call(self, tmp_path):
        """A fileAttachment with null contentBytes AND size==0 must not call $value."""
        client = AsyncMock()
        entries = [_file_entry("empty.bin", content_bytes=None, size=0, aid="Z1")]
        await download_attachments(client, "/me/messages/M1", entries, str(tmp_path))
        client.get_content.assert_not_awaited()
        assert entries[0]["local_path"] is None
        assert entries[0]["note"] is not None
