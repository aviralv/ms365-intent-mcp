import pytest
from unittest.mock import AsyncMock

from ms365_intent_mcp.composers.attachments import (
    body_has_cid,
    classify_attachment,
    enumerate_attachments,
    safe_filename,
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
