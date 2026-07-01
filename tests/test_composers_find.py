"""Tests for find composer."""

from unittest.mock import AsyncMock

import pytest

from ms365_intent_mcp.composers.find import compose_find, _list_user_chats
from ms365_intent_mcp.graph import GraphAPIError
from ms365_intent_mcp.permissions import PermissionRegistry


@pytest.fixture
def full_permissions():
    return PermissionRegistry(["Mail.Read", "Files.Read", "Chat.ReadWrite", "Sites.Read.All"])


class TestFindSearch:
    @pytest.mark.asyncio
    async def test_returns_search_results(self, full_permissions):
        client = AsyncMock()
        client.post = AsyncMock(return_value=_mock_search_response("Q2 Budget"))

        result = await compose_find(
            client=client,
            permissions=full_permissions,
            query="Q2 Budget",
            search_type=None,
        )
        assert "Q2 Budget" in result

    @pytest.mark.asyncio
    async def test_email_type_hint(self, full_permissions):
        client = AsyncMock()
        client.post = AsyncMock(return_value=_mock_search_response("invoice"))

        result = await compose_find(
            client=client,
            permissions=full_permissions,
            query="invoice",
            search_type="email",
        )
        client.post.assert_called_once()
        call_args = client.post.call_args
        payload = call_args[0][1]
        assert "message" in payload["requests"][0]["entityTypes"]

    @pytest.mark.asyncio
    async def test_403_on_chat_message_falls_back(self, full_permissions):
        client = AsyncMock()

        async def _post(endpoint, json_data, headers=None):
            if "chatMessage" in str(json_data):
                raise GraphAPIError(403, "Forbidden", "Chat search not allowed")
            return _mock_search_response("hello")

        client.post = AsyncMock(side_effect=_post)
        result = await compose_find(
            client=client,
            permissions=full_permissions,
            query="hello",
            search_type="message",
        )
        assert result is not None

    @pytest.mark.asyncio
    async def test_no_results(self, full_permissions):
        client = AsyncMock()
        client.post = AsyncMock(return_value={"value": [{"hitsContainers": [{"hits": [], "total": 0}]}]})

        result = await compose_find(
            client=client,
            permissions=full_permissions,
            query="xyzzy_nonexistent",
            search_type=None,
        )
        assert "No results" in result or "xyzzy_nonexistent" in result


class TestListUserChats:
    @pytest.mark.asyncio
    async def test_returns_chats_sorted_by_recency(self):
        client = AsyncMock()
        client.get = AsyncMock(return_value={
            "value": [
                {
                    "id": "chat-old",
                    "members": [],
                    "lastMessagePreview": {"createdDateTime": "2026-06-01T10:00:00Z"},
                },
                {
                    "id": "chat-new",
                    "members": [],
                    "lastMessagePreview": {"createdDateTime": "2026-06-30T10:00:00Z"},
                },
            ]
        })
        chats = await _list_user_chats(client)
        assert [c["id"] for c in chats] == ["chat-new", "chat-old"]

    @pytest.mark.asyncio
    async def test_returns_empty_on_graph_error(self):
        client = AsyncMock()
        client.get = AsyncMock(side_effect=GraphAPIError(403, "Forbidden", "no scope"))
        chats = await _list_user_chats(client)
        assert chats == []


def _mock_search_response(query: str) -> dict:
    return {
        "value": [
            {
                "hitsContainers": [
                    {
                        "hits": [
                            {
                                "hitId": "msg-1",
                                "rank": 1,
                                "summary": f"...{query} related content...",
                                "resource": {
                                    "@odata.type": "#microsoft.graph.message",
                                    "subject": f"{query} Review",
                                    "from": {"emailAddress": {"name": "Alice", "address": "alice@example.com"}},
                                    "receivedDateTime": "2026-05-10T09:00:00Z",
                                    "bodyPreview": f"Please review the {query} document attached.",
                                },
                            }
                        ],
                        "total": 1,
                        "moreResultsAvailable": False,
                    }
                ]
            }
        ]
    }
