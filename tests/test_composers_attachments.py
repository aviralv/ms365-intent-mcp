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
