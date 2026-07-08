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
