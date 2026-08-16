"""Tests for compose composer (discriminated dispatch)."""

from unittest.mock import AsyncMock

import pytest

from ms365_intent_mcp.composers.compose import compose_action, ComposeType
from ms365_intent_mcp.permissions import PermissionRegistry


@pytest.fixture
def full_permissions():
    return PermissionRegistry(["Mail.ReadWrite", "Calendars.ReadWrite", "ChatMessage.Send"])


class TestComposeEmailDraft:
    @pytest.mark.asyncio
    async def test_creates_draft(self, full_permissions):
        client = AsyncMock()
        client.post = AsyncMock(return_value={
            "subject": "Hello",
            "id": "draft-1",
            "toRecipients": [{"emailAddress": {"name": "Bob", "address": "bob@example.com"}}],
        })

        result = await compose_action(
            client=client,
            permissions=full_permissions,
            action_type=ComposeType.EMAIL_DRAFT,
            params={
                "subject": "Hello",
                "body": "Hi Bob",
                "to": [{"email": "bob@example.com", "name": "Bob"}],
            },
        )
        _, markdown = result
        assert "Draft created" in markdown or "✅" in markdown
        client.post.assert_called_once()


class TestComposeReplyDraft:
    @pytest.mark.asyncio
    async def test_creates_reply(self, full_permissions):
        client = AsyncMock()
        client.post = AsyncMock(return_value={
            "subject": "Re: Hello",
            "id": "draft-2",
            "toRecipients": [{"emailAddress": {"name": "Alice", "address": "alice@example.com"}}],
        })

        result = await compose_action(
            client=client,
            permissions=full_permissions,
            action_type=ComposeType.REPLY_DRAFT,
            params={
                "message_id": "msg-123",
                "body": "Thanks!",
                "reply_all": True,
            },
        )
        _, markdown = result
        assert "✅" in markdown


class TestComposeEvent:
    @pytest.mark.asyncio
    async def test_creates_event(self, full_permissions):
        client = AsyncMock()
        client.post = AsyncMock(return_value={
            "subject": "Sync",
            "start": {"dateTime": "2026-05-16T10:00:00"},
            "end": {"dateTime": "2026-05-16T10:30:00"},
            "isOnlineMeeting": False,
            "onlineMeeting": None,
        })

        result = await compose_action(
            client=client,
            permissions=full_permissions,
            action_type=ComposeType.EVENT,
            params={
                "subject": "Sync",
                "start": "2026-05-16T10:00:00",
                "end": "2026-05-16T10:30:00",
                "timezone": "Europe/Berlin",
            },
        )
        _, markdown = result
        assert "✅" in markdown


class TestComposeTeamsMessage:
    @pytest.mark.asyncio
    async def test_sends_message(self, full_permissions):
        client = AsyncMock()
        client.post = AsyncMock(return_value={"id": "msg-456"})

        result = await compose_action(
            client=client,
            permissions=full_permissions,
            action_type=ComposeType.TEAMS_MESSAGE,
            params={
                "chat_id": "19:abc123",
                "content": "Hey team!",
            },
        )
        _, markdown = result
        assert "✅" in markdown or "sent" in markdown.lower()

    @pytest.mark.asyncio
    async def test_sends_message_with_reference_attachments(self, full_permissions):
        client = AsyncMock()
        client.post = AsyncMock(return_value={"id": "msg-789"})

        data, markdown = await compose_action(
            client=client,
            permissions=full_permissions,
            action_type=ComposeType.TEAMS_MESSAGE,
            params={
                "chat_id": "19:abc123",
                "content": "See attached",
                "attachments": [
                    {"name": "report.xlsx", "url": "https://sp.com/sites/team/report.xlsx"},
                    {"name": "notes.pdf", "url": "https://sp.com/sites/team/notes.pdf"},
                ],
            },
        )
        assert "attachment" in markdown.lower()
        # Verify Graph payload shape
        call_payload = client.post.call_args.args[1]
        assert call_payload["body"]["contentType"] == "html"
        assert "attachments" in call_payload
        assert len(call_payload["attachments"]) == 2
        assert call_payload["attachments"][0]["contentType"] == "reference"
        assert call_payload["attachments"][0]["name"] == "report.xlsx"
        assert call_payload["attachments"][0]["contentUrl"] == "https://sp.com/sites/team/report.xlsx"
        # Attachment tags in body
        assert '<attachment id="file-0"></attachment>' in call_payload["body"]["content"]
        assert '<attachment id="file-1"></attachment>' in call_payload["body"]["content"]

    @pytest.mark.asyncio
    async def test_text_content_escaped_when_attachments_force_html(self, full_permissions):
        client = AsyncMock()
        client.post = AsyncMock(return_value={"id": "msg-101"})

        await compose_action(
            client=client,
            permissions=full_permissions,
            action_type=ComposeType.TEAMS_MESSAGE,
            params={
                "chat_id": "19:abc123",
                "content": "A < B & C > D",
                "content_type": "text",
                "attachments": [
                    {"name": "file.txt", "url": "https://sp.com/file.txt"},
                ],
            },
        )
        call_payload = client.post.call_args.args[1]
        body_content = call_payload["body"]["content"]
        assert "&lt;" in body_content
        assert "&amp;" in body_content
        assert "&gt;" in body_content

    @pytest.mark.asyncio
    async def test_no_attachments_sends_plain_text(self, full_permissions):
        client = AsyncMock()
        client.post = AsyncMock(return_value={"id": "msg-200"})

        await compose_action(
            client=client,
            permissions=full_permissions,
            action_type=ComposeType.TEAMS_MESSAGE,
            params={
                "chat_id": "19:abc123",
                "content": "plain text",
                "content_type": "text",
            },
        )
        call_payload = client.post.call_args.args[1]
        assert call_payload["body"]["contentType"] == "text"
        assert "attachments" not in call_payload


class TestComposeEmailForward:
    @pytest.mark.asyncio
    async def test_creates_forward_draft(self, full_permissions):
        client = AsyncMock()
        client.post = AsyncMock(return_value={
            "subject": "FW: Hello",
            "id": "draft-9",
            "toRecipients": [{"emailAddress": {"name": "Carol", "address": "carol@example.com"}}],
        })

        result = await compose_action(
            client=client,
            permissions=full_permissions,
            action_type=ComposeType.EMAIL_FORWARD,
            params={
                "message_id": "msg-777",
                "body": "See <below> & above",
                "to": [{"email": "carol@example.com", "name": "Carol"}],
            },
        )
        _, markdown = result
        assert "✅" in markdown
        endpoint = client.post.call_args.args[0]
        body = client.post.call_args.args[1]
        assert endpoint == "/me/messages/msg-777/createForward"
        assert body["message"]["toRecipients"][0]["emailAddress"]["address"] == "carol@example.com"
        assert body["message"]["body"]["content"] == "See &lt;below&gt; &amp; above"


class TestComposeEventForward:
    @pytest.mark.asyncio
    async def test_forwards_event(self, full_permissions):
        client = AsyncMock()
        client.post = AsyncMock(return_value={})  # 202, empty body

        result = await compose_action(
            client=client,
            permissions=full_permissions,
            action_type=ComposeType.EVENT_FORWARD,
            params={
                "event_id": "AAMkEVT",
                "to": [{"email": "dana@contoso.com", "name": "Dana Swope"}],
                "comment": "Hope you can make it",
            },
        )
        data, markdown = result
        endpoint = client.post.call_args.args[0]
        body = client.post.call_args.args[1]
        assert endpoint == "/me/events/AAMkEVT/forward"
        assert body["ToRecipients"][0]["EmailAddress"]["Address"] == "dana@contoso.com"
        assert body["ToRecipients"][0]["EmailAddress"]["Name"] == "Dana Swope"
        assert body["Comment"] == "Hope you can make it"
        assert "✅" in markdown
        assert data["to"][0]["email"] == "dana@contoso.com"


class TestEndpointIdEncoding:
    """Graph message/event IDs from OWA are standard base64 ('/', '+', '=').
    They must be normalized to Outlook's URL-safe form ('/'→'-', '+'→'_') — the
    form Graph's REST paths accept — before percent-encoding into the path.
    Percent-encoding a raw '/' to '%2F' 404s (RequestBroker ParseUri). Only the
    trailing '=' padding survives normalization and gets percent-encoded.
    """

    RAW_ID = "AAMk/abc+def=="
    ENC_ID = "AAMk-abc_def%3D%3D"

    @pytest.mark.asyncio
    async def test_reply_encodes_message_id(self, full_permissions):
        client = AsyncMock()
        client.post = AsyncMock(return_value={"id": "d", "subject": "Re: x", "toRecipients": []})
        await compose_action(
            client=client,
            permissions=full_permissions,
            action_type=ComposeType.REPLY_DRAFT,
            params={"message_id": self.RAW_ID, "body": "hi", "reply_all": True},
        )
        assert client.post.call_args.args[0] == f"/me/messages/{self.ENC_ID}/createReplyAll"

    @pytest.mark.asyncio
    async def test_email_forward_encodes_message_id(self, full_permissions):
        client = AsyncMock()
        client.post = AsyncMock(return_value={"id": "d", "subject": "FW: x", "toRecipients": []})
        await compose_action(
            client=client,
            permissions=full_permissions,
            action_type=ComposeType.EMAIL_FORWARD,
            params={"message_id": self.RAW_ID, "body": "hi", "to": [{"email": "a@b.com"}]},
        )
        assert client.post.call_args.args[0] == f"/me/messages/{self.ENC_ID}/createForward"

    @pytest.mark.asyncio
    async def test_event_forward_encodes_event_id(self, full_permissions):
        client = AsyncMock()
        client.post = AsyncMock(return_value={})
        await compose_action(
            client=client,
            permissions=full_permissions,
            action_type=ComposeType.EVENT_FORWARD,
            params={"event_id": self.RAW_ID, "to": [{"email": "a@b.com"}]},
        )
        assert client.post.call_args.args[0] == f"/me/events/{self.ENC_ID}/forward"


class TestComposeMissingPermission:
    @pytest.mark.asyncio
    async def test_no_mail_scope(self):
        client = AsyncMock()
        permissions = PermissionRegistry(["Calendars.ReadWrite"])

        result = await compose_action(
            client=client,
            permissions=permissions,
            action_type=ComposeType.EMAIL_DRAFT,
            params={"subject": "Test", "body": "Hi", "to": [{"email": "x@x.com"}]},
        )
        _, markdown = result
        assert "Mail.ReadWrite" in markdown
