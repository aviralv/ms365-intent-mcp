"""Tests for people composer."""

from unittest.mock import AsyncMock

import pytest

from ms365_intent_mcp.composers.people import compose_people
from ms365_intent_mcp.graph import GraphAPIError
from ms365_intent_mcp.permissions import PermissionRegistry


@pytest.fixture
def full_permissions():
    return PermissionRegistry(["People.Read", "Mail.Read", "Chat.ReadWrite"])


@pytest.fixture
def no_people_permissions():
    return PermissionRegistry(["Contacts.Read", "Mail.Read"])


class TestPeopleBasic:
    @pytest.mark.asyncio
    async def test_returns_person_info(self, full_permissions):
        client = AsyncMock()

        async def _get(endpoint, params=None, headers=None):
            if "/me/people" in endpoint:
                return _mock_people_response()
            if "/me/messages" in endpoint:
                return {"value": []}
            if "/me/chats" in endpoint:
                return {"value": []}
            return {"value": []}

        client.get = AsyncMock(side_effect=_get)
        result = await compose_people(client=client, permissions=full_permissions, query="alice")
        assert "Alice Smith" in result

    @pytest.mark.asyncio
    async def test_no_people_read_uses_contacts_fallback(self, no_people_permissions):
        client = AsyncMock()

        async def _get(endpoint, params=None, headers=None):
            if "/me/contacts" in endpoint:
                return {"value": [{"displayName": "Alice Contact", "emailAddresses": [{"address": "alice@example.com"}]}]}
            return {"value": []}

        client.get = AsyncMock(side_effect=_get)
        result = await compose_people(client=client, permissions=no_people_permissions, query="alice")
        assert "Alice" in result

    @pytest.mark.asyncio
    async def test_no_results_returns_not_found(self, full_permissions):
        client = AsyncMock()
        client.get = AsyncMock(return_value={"value": []})

        result = await compose_people(client=client, permissions=full_permissions, query="zzz_nobody")
        assert "No results" in result or "zzz_nobody" in result

    @pytest.mark.asyncio
    async def test_mail_failure_still_returns_person(self, full_permissions):
        client = AsyncMock()

        async def _get(endpoint, params=None, headers=None):
            if "/me/people" in endpoint:
                return _mock_people_response()
            if "/me/messages" in endpoint:
                raise GraphAPIError(500, "InternalError", "mail down")
            return {"value": []}

        client.get = AsyncMock(side_effect=_get)
        result = await compose_people(client=client, permissions=full_permissions, query="alice")
        assert "Alice Smith" in result


def _mock_people_response():
    return {
        "value": [
            {
                "displayName": "Alice Smith",
                "jobTitle": "Engineer",
                "emailAddresses": [{"address": "alice@example.com"}],
            }
        ]
    }


class TestFindChatWithPerson:
    def test_email_match_wins(self):
        from ms365_intent_mcp.composers.people import _find_chat_with_person
        chats = [
            {"id": "1", "members": [{"displayName": "Bob", "email": "bob@example.com"}]},
            {"id": "2", "members": [{"displayName": "Alice Smith", "email": "alice@example.com"}]},
        ]
        result = _find_chat_with_person(chats, "Alice Smith", "alice@example.com")
        assert result["id"] == "2"

    def test_email_match_preferred_over_name(self):
        from ms365_intent_mcp.composers.people import _find_chat_with_person
        chats = [
            {"id": "name-match", "members": [{"displayName": "Alice Smith", "email": "other@example.com"}]},
            {"id": "email-match", "members": [{"displayName": "AS", "email": "alice@example.com"}]},
        ]
        result = _find_chat_with_person(chats, "Alice Smith", "alice@example.com")
        assert result["id"] == "email-match"

    def test_avi_does_not_match_aviral(self):
        from ms365_intent_mcp.composers.people import _find_chat_with_person
        chats = [
            {"id": "1", "members": [{"displayName": "Aviral Patel", "email": ""}]},
        ]
        result = _find_chat_with_person(chats, "Avi", "")
        assert result is None

    def test_full_name_matches_when_no_email(self):
        from ms365_intent_mcp.composers.people import _find_chat_with_person
        chats = [
            {"id": "1", "members": [{"displayName": "Alice Smith", "email": ""}]},
        ]
        result = _find_chat_with_person(chats, "Alice Smith", "")
        assert result["id"] == "1"

    def test_no_match_returns_none(self):
        from ms365_intent_mcp.composers.people import _find_chat_with_person
        chats = [
            {"id": "1", "members": [{"displayName": "Bob Jones", "email": "bob@example.com"}]},
        ]
        result = _find_chat_with_person(chats, "Alice Smith", "alice@example.com")
        assert result is None

    def test_empty_target_returns_none(self):
        from ms365_intent_mcp.composers.people import _find_chat_with_person
        chats = [{"id": "1", "members": [{"displayName": "Alice", "email": "a@b.com"}]}]
        assert _find_chat_with_person(chats, "", "") is None
