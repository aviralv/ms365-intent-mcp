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
        _, result = await compose_people(client=client, permissions=full_permissions, query="alice")
        assert "Alice Smith" in result

    @pytest.mark.asyncio
    async def test_no_people_read_uses_contacts_fallback(self, no_people_permissions):
        client = AsyncMock()

        async def _get(endpoint, params=None, headers=None):
            if "/me/contacts" in endpoint:
                return {"value": [{"displayName": "Alice Contact", "emailAddresses": [{"address": "alice@example.com"}]}]}
            return {"value": []}

        client.get = AsyncMock(side_effect=_get)
        _, result = await compose_people(client=client, permissions=no_people_permissions, query="alice")
        assert "Alice" in result

    @pytest.mark.asyncio
    async def test_no_results_returns_not_found(self, full_permissions):
        client = AsyncMock()
        client.get = AsyncMock(return_value={"value": []})

        _, result = await compose_people(client=client, permissions=full_permissions, query="zzz_nobody")
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
        _, result = await compose_people(client=client, permissions=full_permissions, query="alice")
        assert "Alice Smith" in result


class TestPeopleChatUrl:
    """#37 Option 2: when people() finds a 1:1 chat with the person, it must
    surface that chat's thread URL (webUrl) so the caller can render it via
    resolve() — not just the last-message preview."""

    _CHAT_URL = "https://teams.microsoft.com/l/chat/19:alice@unq.gbl.spaces/0"

    def _client(self):
        client = AsyncMock()

        async def _get(endpoint, params=None, headers=None):
            if "/me/people" in endpoint:
                return _mock_people_response()
            if "/me/messages" in endpoint:
                return {"value": []}
            if "/me/chats" in endpoint:
                return {
                    "value": [
                        {
                            "id": "chat-alice",
                            "webUrl": self._CHAT_URL,
                            "members": [{"displayName": "Alice Smith", "email": "alice@example.com"}],
                            "lastMessagePreview": {
                                "body": {"content": "<p>hi there</p>"},
                                "createdDateTime": "2026-06-30T10:00:00Z",
                            },
                        }
                    ]
                }
            return {"value": []}

        client.get = AsyncMock(side_effect=_get)
        return client

    @pytest.mark.asyncio
    async def test_structured_recent_chat_includes_chat_id_and_url(self, full_permissions):
        data, _ = await compose_people(client=self._client(), permissions=full_permissions, query="alice")
        assert data["recent_chat"] is not None
        assert data["recent_chat"]["chat_id"] == "chat-alice"
        assert data["recent_chat"]["chat_url"] == self._CHAT_URL

    @pytest.mark.asyncio
    async def test_markdown_includes_open_chat_link(self, full_permissions):
        _, result = await compose_people(client=self._client(), permissions=full_permissions, query="alice")
        assert "open chat" in result
        assert self._CHAT_URL in result


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


class TestLookupPersonViaChats:
    def test_synthesizes_person_from_chat_member(self):
        from ms365_intent_mcp.composers.people import _lookup_person_via_chats
        chats = [
            {
                "id": "chat-y",
                "webUrl": "https://teams.microsoft.com/l/chat/19:yev/0",
                "members": [
                    {"displayName": "Me", "userId": "me-id"},
                    {"displayName": "Yevhen Kushnirenko", "userId": "yev-id",
                     "email": "yevhen.k@sap.com"},
                ],
                "lastMessagePreview": {"createdDateTime": "2026-07-16T09:00:00Z"},
            }
        ]
        people = _lookup_person_via_chats(chats, "Yevhen", me_id="me-id")
        assert len(people) == 1
        assert people[0]["displayName"] == "Yevhen Kushnirenko"
        assert people[0]["emailAddresses"] == [{"address": "yevhen.k@sap.com"}]
        assert people[0]["_source_chat"]["id"] == "chat-y"

    def test_self_excluded_by_user_id(self):
        from ms365_intent_mcp.composers.people import _lookup_person_via_chats
        chats = [
            {"id": "g", "members": [
                {"displayName": "Aviral Vaid", "userId": "me-id"},
                {"displayName": "Aviral Kumar", "userId": "other-id"},
            ]}
        ]
        people = _lookup_person_via_chats(chats, "Aviral", me_id="me-id")
        assert [p["displayName"] for p in people] == ["Aviral Kumar"]

    def test_avi_does_not_match_aviral(self):
        from ms365_intent_mcp.composers.people import _lookup_person_via_chats
        chats = [{"id": "1", "members": [{"displayName": "Aviral Patel", "userId": "x"}]}]
        assert _lookup_person_via_chats(chats, "Avi", me_id="") == []

    def test_dedup_prefers_user_id_then_email_then_name(self):
        from ms365_intent_mcp.composers.people import _lookup_person_via_chats
        chats = [
            {"id": "a", "members": [{"displayName": "Sam Lee", "userId": "sam"}]},
            {"id": "b", "members": [{"displayName": "Sam Lee", "userId": "sam"}]},
            {"id": "c", "members": [{"displayName": "Sam Lee", "email": "sam@x.com"}]},
        ]
        people = _lookup_person_via_chats(chats, "Sam Lee", me_id="")
        # userId "sam" dedups a+b to one; email-only "sam@x.com" is a distinct key
        assert len(people) == 2
        assert people[0]["_source_chat"]["id"] == "a"

    def test_multiple_distinct_people_ordered_by_recency(self):
        from ms365_intent_mcp.composers.people import _lookup_person_via_chats
        chats = [
            {"id": "recent", "members": [{"displayName": "Dana First", "userId": "d1"}]},
            {"id": "older", "members": [{"displayName": "Dana Second", "userId": "d2"}]},
        ]
        people = _lookup_person_via_chats(chats, "Dana", me_id="")
        assert [p["displayName"] for p in people] == ["Dana First", "Dana Second"]

    def test_member_without_userid_or_email_deduped_by_name(self):
        from ms365_intent_mcp.composers.people import _lookup_person_via_chats
        chats = [
            {"id": "a", "members": [{"displayName": "Guest Person"}]},
            {"id": "b", "members": [{"displayName": "Guest Person"}]},
        ]
        people = _lookup_person_via_chats(chats, "Guest Person", me_id="")
        assert len(people) == 1
        assert people[0]["emailAddresses"] == []

    def test_member_with_no_synthesizable_identity_skipped(self):
        from ms365_intent_mcp.composers.people import _lookup_person_via_chats
        chats = [{"id": "a", "members": [{"displayName": ""}]}]
        assert _lookup_person_via_chats(chats, "anything", me_id="") == []


class TestPeopleChatFallback:
    def _client(self, chats_value, count_holder=None):
        client = AsyncMock()

        async def _get(endpoint, params=None, headers=None):
            if "/me/people" in endpoint:
                return {"value": []}
            if "/me/contacts" in endpoint:
                return {"value": []}
            if endpoint == "/me":
                return {"id": "me-id"}
            if "/me/messages" in endpoint:
                return {"value": []}
            if "/me/chats" in endpoint:
                if count_holder is not None:
                    count_holder["chats"] = count_holder.get("chats", 0) + 1
                return {"value": chats_value}
            return {"value": []}

        client.get = AsyncMock(side_effect=_get)
        return client

    _YEV_CHAT = {
        "id": "chat-yev",
        "webUrl": "https://teams.microsoft.com/l/chat/19:yev@unq.gbl.spaces/0",
        "members": [
            {"displayName": "Me", "userId": "me-id"},
            {"displayName": "Yevhen Kushnirenko", "userId": "yev-id",
             "email": "yevhen.k@sap.com"},
        ],
        "lastMessagePreview": {
            "body": {"content": "<p>ping</p>"},
            "createdDateTime": "2026-07-16T09:00:00Z",
        },
    }

    @pytest.mark.asyncio
    async def test_yevhen_resolved_via_chat(self, full_permissions):
        client = self._client([self._YEV_CHAT])
        data, markdown = await compose_people(
            client=client, permissions=full_permissions, query="Yevhen")
        assert "Yevhen Kushnirenko" in markdown
        assert data["recent_chat"] is not None
        assert data["recent_chat"]["chat_url"] == self._YEV_CHAT["webUrl"]

    @pytest.mark.asyncio
    async def test_chats_fetched_exactly_once(self, full_permissions):
        counter: dict = {}
        client = self._client([self._YEV_CHAT], count_holder=counter)
        await compose_people(client=client, permissions=full_permissions, query="Yevhen")
        assert counter["chats"] == 1

    @pytest.mark.asyncio
    async def test_no_chat_read_scope_returns_not_found(self):
        perms = PermissionRegistry(["People.Read", "Mail.Read"])
        client = self._client([self._YEV_CHAT])
        _, markdown = await compose_people(
            client=client, permissions=perms, query="Yevhen")
        assert "No results" in markdown

    @pytest.mark.asyncio
    async def test_chat_enumeration_error_returns_not_found(self, full_permissions):
        client = AsyncMock()

        async def _get(endpoint, params=None, headers=None):
            if "/me/chats" in endpoint:
                raise GraphAPIError(500, "InternalError", "chats down")
            return {"value": []}

        client.get = AsyncMock(side_effect=_get)
        _, markdown = await compose_people(
            client=client, permissions=full_permissions, query="Yevhen")
        assert "No results" in markdown

    @pytest.mark.asyncio
    async def test_people_hit_still_works_and_fetches_chats_once(self, full_permissions):
        counter: dict = {}
        client = AsyncMock()

        async def _get(endpoint, params=None, headers=None):
            if "/me/people" in endpoint:
                return _mock_people_response()
            if "/me/messages" in endpoint:
                return {"value": []}
            if endpoint == "/me":
                return {"id": "me-id"}
            if "/me/chats" in endpoint:
                counter["chats"] = counter.get("chats", 0) + 1
                return {"value": []}
            return {"value": []}

        client.get = AsyncMock(side_effect=_get)
        _, markdown = await compose_people(
            client=client, permissions=full_permissions, query="alice")
        assert "Alice Smith" in markdown
        assert counter["chats"] == 1

    @pytest.mark.asyncio
    async def test_me_identity_failure_still_resolves_via_fallback(self, full_permissions):
        client = AsyncMock()

        async def _get(endpoint, params=None, headers=None):
            if "/me/people" in endpoint:
                return {"value": []}
            if "/me/contacts" in endpoint:
                return {"value": []}
            if endpoint == "/me":
                raise TimeoutError("me timed out")
            if "/me/messages" in endpoint:
                return {"value": []}
            if "/me/chats" in endpoint:
                return {"value": [self._YEV_CHAT]}
            return {"value": []}

        client.get = AsyncMock(side_effect=_get)
        _, markdown = await compose_people(
            client=client, permissions=full_permissions, query="Yevhen")
        assert "Yevhen Kushnirenko" in markdown


class TestLookupQueryEscaping:
    @pytest.mark.asyncio
    async def test_contacts_query_is_escaped(self):
        from ms365_intent_mcp.composers.people import _lookup_person
        client = AsyncMock()
        captured = {}

        async def _get(endpoint, params=None, headers=None):
            if "/me/contacts" in endpoint:
                captured["search"] = params["$search"]
            return {"value": []}

        client.get = AsyncMock(side_effect=_get)
        perms = PermissionRegistry(["Contacts.Read"])  # no People.Read → skip /me/people
        await _lookup_person(client, perms, "O'Malley")
        assert "O''Malley" in captured["search"]

    @pytest.mark.asyncio
    async def test_people_select_uses_scored_email_addresses(self):
        from ms365_intent_mcp.composers.people import _lookup_person
        client = AsyncMock()
        captured = {}

        async def _get(endpoint, params=None, headers=None):
            if "/me/people" in endpoint:
                captured["select"] = params["$select"]
            return {"value": []}

        client.get = AsyncMock(side_effect=_get)
        perms = PermissionRegistry(["People.Read"])
        await _lookup_person(client, perms, "alice")
        assert "scoredEmailAddresses" in captured["select"]
        assert "emailAddresses" not in captured["select"].replace("scoredEmailAddresses", "")


class TestUsersDirectoryTier:
    @pytest.mark.asyncio
    async def test_users_tier_resolves_when_earlier_tiers_email_less(self):
        client = AsyncMock()

        async def _get(endpoint, params=None, headers=None):
            if "/me/people" in endpoint:
                return {"value": []}
            if "/me/contacts" in endpoint:
                return {"value": []}
            if "/users" in endpoint:
                assert headers and headers.get("ConsistencyLevel") == "eventual"
                return {"value": [{
                    "displayName": "Karlbowski, Marcus",
                    "mail": "marcus.karlbowski@sap.com",
                    "userPrincipalName": "marcus.karlbowski@sap.com",
                }]}
            return {"value": []}

        client.get = AsyncMock(side_effect=_get)
        perms = PermissionRegistry(["Contacts.Read", "User.ReadBasic.All", "Mail.Read"])
        data, _ = await compose_people(client, perms, "Marcus Karlbowski")
        assert data["email"] == "marcus.karlbowski@sap.com"

    @pytest.mark.asyncio
    async def test_users_tier_skipped_without_scope(self):
        client = AsyncMock()
        calls = []

        async def _get(endpoint, params=None, headers=None):
            calls.append(endpoint)
            return {"value": []}

        client.get = AsyncMock(side_effect=_get)
        perms = PermissionRegistry(["Contacts.Read"])  # no User.ReadBasic.All
        data, _ = await compose_people(client, perms, "Marcus")
        assert not any("/users" in c for c in calls)
        assert data["email"] == ""  # graceful: no error

    @pytest.mark.asyncio
    async def test_multiple_hits_withholds_email(self):
        client = AsyncMock()

        async def _get(endpoint, params=None, headers=None):
            if "/users" in endpoint:
                return {"value": [
                    {"displayName": "Marcus Kern", "mail": "marcus.kern@sap.com"},
                    {"displayName": "Marcus Nebel", "mail": "marcus.nebel@sap.com"},
                ]}
            return {"value": []}

        client.get = AsyncMock(side_effect=_get)
        perms = PermissionRegistry(["Contacts.Read", "User.ReadBasic.All"])
        data, markdown = await compose_people(client, perms, "Marcus")
        assert data["email"] == ""            # withheld — ambiguous
        assert data["name"] == "Marcus Kern"  # top hit name still populated
        assert "Marcus Nebel" in markdown     # candidates surfaced

    @pytest.mark.asyncio
    async def test_single_hit_with_email_is_confident(self):
        client = AsyncMock()

        async def _get(endpoint, params=None, headers=None):
            if "/users" in endpoint:
                return {"value": [{"displayName": "Karlbowski, Marcus", "mail": "marcus.karlbowski@sap.com"}]}
            return {"value": []}

        client.get = AsyncMock(side_effect=_get)
        perms = PermissionRegistry(["Contacts.Read", "User.ReadBasic.All"])
        data, _ = await compose_people(client, perms, "Marcus Karlbowski")
        assert data["email"] == "marcus.karlbowski@sap.com"


class TestExtractEmail:
    def test_contact_shape(self):
        from ms365_intent_mcp.composers.people import _extract_email
        assert _extract_email({"emailAddresses": [{"address": "a@b.com"}]}) == "a@b.com"

    def test_person_scored_shape(self):
        from ms365_intent_mcp.composers.people import _extract_email
        assert _extract_email({"scoredEmailAddresses": [{"address": "p@b.com"}]}) == "p@b.com"

    def test_users_mail_shape(self):
        from ms365_intent_mcp.composers.people import _extract_email
        assert _extract_email({"mail": "u@b.com", "userPrincipalName": "u@b.com"}) == "u@b.com"

    def test_users_upn_fallback_when_mail_null(self):
        from ms365_intent_mcp.composers.people import _extract_email
        assert _extract_email({"mail": None, "userPrincipalName": "u@sap.com"}) == "u@sap.com"

    def test_guest_upn_rejected(self):
        from ms365_intent_mcp.composers.people import _extract_email
        assert _extract_email({"mail": None, "userPrincipalName": "x_ext.com#EXT#@t.onmicrosoft.com"}) == ""

    def test_none_when_empty(self):
        from ms365_intent_mcp.composers.people import _extract_email
        assert _extract_email({"displayName": "No Email"}) == ""
