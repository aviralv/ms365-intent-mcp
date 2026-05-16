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
                return {"value": [{"displayName": "Alice Contact", "emailAddresses": [{"address": "alice@sap.com"}]}]}
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
                "emailAddresses": [{"address": "alice@sap.com"}],
            }
        ]
    }
