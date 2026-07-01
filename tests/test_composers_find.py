"""Tests for find composer."""

from unittest.mock import AsyncMock

import pytest

from ms365_intent_mcp.composers.find import compose_find, _list_user_chats, _fetch_chat_messages
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
    async def test_message_type_uses_chat_enumeration_not_search(self, full_permissions):
        client = AsyncMock()
        client.get = AsyncMock(return_value={"value": []})
        client.post = AsyncMock()

        result = await compose_find(
            client=client,
            permissions=full_permissions,
            query="hello",
            search_type="message",
        )
        client.post.assert_not_called()
        client.get.assert_called()
        assert "No results" in result

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
    async def test_raises_graph_error(self):
        client = AsyncMock()
        client.get = AsyncMock(side_effect=GraphAPIError(403, "Forbidden", "no scope"))
        with pytest.raises(GraphAPIError):
            await _list_user_chats(client)


class TestFetchChatMessages:
    @pytest.mark.asyncio
    async def test_filters_by_query_substring_case_insensitive(self):
        client = AsyncMock()
        client.get = AsyncMock(return_value={
            "value": [
                {"id": "m1", "body": {"content": "<p>Second Brain rocks</p>"}, "from": {"user": {"displayName": "Diana"}}, "createdDateTime": "2026-06-30T10:00:00Z"},
                {"id": "m2", "body": {"content": "<p>Unrelated</p>"}, "from": {"user": {"displayName": "Diana"}}, "createdDateTime": "2026-06-29T10:00:00Z"},
                {"id": "m3", "body": {"content": "<p>second brain again</p>"}, "from": {"user": {"displayName": "Diana"}}, "createdDateTime": "2026-06-28T10:00:00Z"},
            ]
        })
        hits = await _fetch_chat_messages(client, "chat-1", ["second brain"])
        assert [h["id"] for h in hits] == ["m1", "m3"]
        assert all(h["_chat_id"] == "chat-1" for h in hits)

    @pytest.mark.asyncio
    async def test_returns_empty_on_graph_error(self):
        client = AsyncMock()
        client.get = AsyncMock(side_effect=GraphAPIError(403, "Forbidden", "no scope"))
        hits = await _fetch_chat_messages(client, "chat-1", ["anything"])
        assert hits == []

    @pytest.mark.asyncio
    async def test_strips_html_before_matching(self):
        client = AsyncMock()
        client.get = AsyncMock(return_value={
            "value": [
                {"id": "m1", "body": {"content": '<div><at id="0">Diana</at> found it</div>'}, "from": {"user": {"displayName": "X"}}, "createdDateTime": "2026-06-30T10:00:00Z"},
            ]
        })
        hits = await _fetch_chat_messages(client, "chat-1", ["Diana found"])
        assert len(hits) == 1

    @pytest.mark.asyncio
    async def test_empty_needles_returns_empty(self):
        client = AsyncMock()
        client.get = AsyncMock(return_value={
            "value": [
                {"id": "m1", "body": {"content": "<p>anything</p>"}, "from": {"user": {"displayName": "X"}}, "createdDateTime": "2026-06-30T10:00:00Z"},
            ]
        })
        hits = await _fetch_chat_messages(client, "chat-1", ["   "])
        assert hits == []
        client.get.assert_not_called()

    @pytest.mark.asyncio
    async def test_decodes_html_entities_before_matching(self):
        client = AsyncMock()
        client.get = AsyncMock(return_value={
            "value": [
                {"id": "m1", "body": {"content": "<p>Q&amp;A session</p>"}, "from": {"user": {"displayName": "X"}}, "createdDateTime": "2026-06-30T10:00:00Z"},
            ]
        })
        hits = await _fetch_chat_messages(client, "chat-1", ["Q&A"])
        assert len(hits) == 1

    @pytest.mark.asyncio
    async def test_matches_any_needle(self):
        client = AsyncMock()
        client.get = AsyncMock(return_value={
            "value": [
                {"id": "m1", "body": {"content": "<p>second brain note</p>"}, "from": {"user": {"displayName": "X"}}, "createdDateTime": "2026-06-30T10:00:00Z"},
                {"id": "m2", "body": {"content": "<p>brain dump</p>"}, "from": {"user": {"displayName": "X"}}, "createdDateTime": "2026-06-29T10:00:00Z"},
                {"id": "m3", "body": {"content": "<p>unrelated stuff</p>"}, "from": {"user": {"displayName": "X"}}, "createdDateTime": "2026-06-28T10:00:00Z"},
            ]
        })
        hits = await _fetch_chat_messages(client, "chat-1", ["second", "brain"])
        hit_ids = {h["id"] for h in hits}
        assert "m1" in hit_ids
        assert "m2" in hit_ids
        assert "m3" not in hit_ids
        # single Graph call (no per-needle fan-out)
        client.get.assert_called_once()


class TestSearchQueryPermissionErrorSurface:
    @pytest.mark.asyncio
    async def test_channel_message_error_points_at_alternatives(self, full_permissions):
        client = AsyncMock()
        client.post = AsyncMock(side_effect=GraphAPIError(
            403,
            "Forbidden",
            "Access to ChatMessage in Graph API requires: Chat.Read, ChannelMessage.Read.All",
        ))
        from ms365_intent_mcp.composers.find import _search_single
        result = await _search_single(client, "hello", ["message"])
        assert "resolve" in result
        assert "people" in result
        assert "whats_new" in result


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


from ms365_intent_mcp.composers.find import _prefilter_chats_by_query


class TestPrefilterChatsByQuery:
    def test_filters_when_query_matches_member_name(self):
        chats = [
            {"id": "c1", "members": [{"displayName": "Diana Veit"}, {"displayName": "Me"}]},
            {"id": "c2", "members": [{"displayName": "Bob"}, {"displayName": "Me"}]},
            {"id": "c3", "members": [{"displayName": "diana smith"}, {"displayName": "Me"}]},
        ]
        result = _prefilter_chats_by_query(chats, "Diana second brain")
        assert [c["id"] for c in result[0]] == ["c1", "c3"]
        assert "diana" in result[1]

    def test_passthrough_when_no_match(self):
        chats = [
            {"id": "c1", "members": [{"displayName": "Bob"}]},
            {"id": "c2", "members": [{"displayName": "Alice"}]},
        ]
        result = _prefilter_chats_by_query(chats, "roadmap plan")
        assert [c["id"] for c in result[0]] == ["c1", "c2"]
        assert result[1] == set()

    def test_ignores_short_words(self):
        chats = [
            {"id": "c1", "members": [{"displayName": "Ai team"}]},
            {"id": "c2", "members": [{"displayName": "Bob"}]},
        ]
        result = _prefilter_chats_by_query(chats, "ai plan")
        assert [c["id"] for c in result[0]] == ["c1", "c2"]
        assert result[1] == set()

    def test_empty_chats_returns_empty(self):
        assert _prefilter_chats_by_query([], "anything") == ([], set())


class TestSearchChatMessages:
    @pytest.mark.asyncio
    async def test_end_to_end_returns_matching_messages(self, full_permissions):
        client = AsyncMock()

        chats_response = {
            "value": [
                {
                    "id": "chat-diana",
                    "members": [{"displayName": "Diana Veit"}, {"displayName": "Me"}],
                    "lastMessagePreview": {"createdDateTime": "2026-06-30T10:00:00Z"},
                },
                {
                    "id": "chat-bob",
                    "members": [{"displayName": "Bob"}, {"displayName": "Me"}],
                    "lastMessagePreview": {"createdDateTime": "2026-06-29T10:00:00Z"},
                },
            ]
        }
        messages_by_chat = {
            "chat-diana": {
                "value": [
                    {"id": "m1", "body": {"content": "<p>Second brain idea</p>"}, "from": {"user": {"displayName": "Diana"}}, "createdDateTime": "2026-06-30T10:00:00Z"},
                ]
            },
        }

        async def _get(path, params=None):
            if path == "/me/chats":
                return chats_response
            for chat_id, payload in messages_by_chat.items():
                if path == f"/chats/{chat_id}/messages":
                    return payload
            return {"value": []}

        client.get = AsyncMock(side_effect=_get)

        result = await compose_find(
            client=client,
            permissions=full_permissions,
            query="Diana second brain",
            search_type="message",
        )
        assert "Second brain idea" in result
        assert "Bob" not in result

    @pytest.mark.asyncio
    async def test_returns_no_results_when_nothing_matches(self, full_permissions):
        client = AsyncMock()

        async def _get(path, params=None):
            if path == "/me/chats":
                return {"value": [{"id": "c1", "members": [], "lastMessagePreview": {"createdDateTime": "2026-06-30T10:00:00Z"}}]}
            return {"value": [{"id": "m1", "body": {"content": "hello"}, "from": {"user": {"displayName": "X"}}, "createdDateTime": "2026-06-30T10:00:00Z"}]}

        client.get = AsyncMock(side_effect=_get)

        result = await compose_find(
            client=client,
            permissions=full_permissions,
            query="nonexistent phrase xyz",
            search_type="message",
        )
        assert "No results" in result

    @pytest.mark.asyncio
    async def test_person_only_query_returns_recent_messages(self, full_permissions):
        client = AsyncMock()

        chats_response = {
            "value": [
                {
                    "id": "chat-diana",
                    "members": [{"displayName": "Diana Veit"}, {"displayName": "Me"}],
                    "lastMessagePreview": {"createdDateTime": "2026-06-30T10:00:00Z"},
                },
            ]
        }

        async def _get(path, params=None):
            if path == "/me/chats":
                return chats_response
            if path == "/chats/chat-diana/messages":
                return {
                    "value": [
                        {"id": "m1", "body": {"content": "<p>quick note</p>"}, "from": {"user": {"displayName": "Diana"}}, "createdDateTime": "2026-06-30T10:00:00Z"},
                    ]
                }
            return {"value": []}

        client.get = AsyncMock(side_effect=_get)

        result = await compose_find(
            client=client,
            permissions=full_permissions,
            query="Diana",
            search_type="message",
        )
        assert "quick note" in result

    @pytest.mark.asyncio
    async def test_chats_auth_error_surfaces_error(self, full_permissions):
        client = AsyncMock()
        client.get = AsyncMock(side_effect=GraphAPIError(403, "Forbidden", "no scope"))
        result = await compose_find(
            client=client,
            permissions=full_permissions,
            query="anything",
            search_type="message",
        )
        assert "Find unavailable" in result or "unavailable" in result
