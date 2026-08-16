"""Tests for URL resolver dispatch table."""

import re
import urllib.parse

import pytest

from ms365_intent_mcp.resolver import (
    _CHAT_THREAD_URL_BASE,
    _PATTERNS,
    UrlParseError,
    _decode_upn,
    build_chat_thread_url,
    resolve_url,
)


def _ctx(payload: dict) -> str:
    """URL-encode a JSON context dict for use in Teams URLs."""
    import json

    return urllib.parse.quote(json.dumps(payload))


class TestChannelMessage:
    def test_with_group_id(self):
        ctx = _ctx({"groupId": "team-uuid-123", "tid": "tenant-id"})
        url = f"https://teams.microsoft.com/l/message/19:abc@thread.tacv2/1234567890.123456?context={ctx}"
        result = resolve_url(url)
        assert result.url_type == "channel_message"
        assert (
            result.graph_endpoint
            == "/teams/team-uuid-123/channels/19:abc@thread.tacv2/messages/1234567890.123456"
        )
        assert result.extra["group_id"] == "team-uuid-123"

    def test_without_group_id(self):
        ctx = _ctx({"tid": "tenant-id"})
        url = f"https://teams.microsoft.com/l/message/19:abc@thread.tacv2/1234567890.123456?context={ctx}"
        result = resolve_url(url)
        assert result.url_type == "channel_message"
        assert result.graph_endpoint.startswith("/chats/")
        assert "group_id" not in result.extra

    def test_message_id_preserves_dot(self):
        ctx = _ctx({"groupId": "team-uuid-123"})
        url = f"https://teams.microsoft.com/l/message/19:abc@thread.tacv2/1234567890.123456?context={ctx}"
        result = resolve_url(url)
        assert "1234567890.123456" in result.graph_endpoint


class TestChatMessage:
    def test_specific_message_endpoint(self):
        url = "https://teams.microsoft.com/l/message/19:somechat@unq.gbl.spaces/1234567890.123456"
        result = resolve_url(url)
        assert result.url_type == "chat_message"
        assert (
            result.graph_endpoint == "/chats/19:somechat@unq.gbl.spaces/messages/1234567890.123456"
        )

    def test_percent_encoded_chat_message_url_parses(self):
        url = (
            "https://teams.microsoft.com/l/message/19%3Asomechat%40unq.gbl.spaces/1234567890.123456"
        )
        result = resolve_url(url)
        assert result.url_type == "chat_message"
        assert (
            result.graph_endpoint == "/chats/19:somechat@unq.gbl.spaces/messages/1234567890.123456"
        )


class TestMeeting:
    def test_endpoint_is_calendar_view(self):
        ctx = _ctx({"Tid": "tenant-id"})
        url = (
            f"https://teams.microsoft.com/l/meetup-join/19:meeting_abc123@thread.v2/0?context={ctx}"
        )
        result = resolve_url(url)
        assert result.url_type == "meeting"
        assert result.graph_endpoint == "/me/calendarView"

    def test_extra_contains_thread_id(self):
        thread_segment = "19:meeting_abc123@thread.v2"
        url = f"https://teams.microsoft.com/l/meetup-join/{urllib.parse.quote(thread_segment)}/0"
        result = resolve_url(url)
        assert "thread_id" in result.extra
        assert result.extra["thread_id"] == thread_segment


class TestEmail:
    def test_mail_id_in_endpoint(self):
        msg_id = "AAkALgAAAAAAHYQDEapmEc2byACqAC-EWg0AKvpMkzmkWEOzFzAmIBTbDwAAAAAB_LqXAAA="
        url = f"https://outlook.office365.com/mail/id/{msg_id}"
        result = resolve_url(url)
        assert result.url_type == "email"
        # Endpoint percent-encodes reserved chars so Graph doesn't interpret
        # embedded '/' or '=' as path/query separators.
        assert result.graph_endpoint == f"/me/messages/{urllib.parse.quote(msg_id, safe='')}"

    def test_owa_itemid_shape(self):
        """webLink from Graph Search returns owa/?ItemID=... form."""
        msg_id = "AAMkADJiMzkyOWY1LWJkY2ItNGY2ZC05ZTE2LTY3OTQyNDBmYWQ1NQBGAAAAAAB_LqXAAA="
        url = f"https://outlook.office365.com/owa/?ItemID={msg_id}&exvsurl=1&viewmodel=ReadMessageItem"
        result = resolve_url(url)
        assert result.url_type == "email"
        assert result.graph_endpoint == f"/me/messages/{urllib.parse.quote(msg_id, safe='')}"

    def test_owa_itemid_when_not_first_param(self):
        msg_id = "AAMkADJiMzkyOWY1LWJkY2I="
        url = f"https://outlook.office365.com/owa/?exvsurl=1&ItemID={msg_id}&viewmodel=ReadMessageItem"
        result = resolve_url(url)
        assert result.url_type == "email"
        assert result.graph_endpoint == f"/me/messages/{urllib.parse.quote(msg_id, safe='')}"

    def test_owa_itemid_with_embedded_slash_is_normalized(self):
        """Search-returned IDs contain '/' (standard-base64) after URL-decoding.
        Graph accepts only the URL-safe variant, so '/' must be rewritten to
        '-' (Outlook's non-standard base64url mapping — verified against the
        live API against hitId comparisons)."""
        msg_id_encoded = "AAMkADJi%2FzOHQAAA%3D"
        url = f"https://outlook.office365.com/owa/?ItemID={msg_id_encoded}&exvsurl=1"
        result = resolve_url(url)
        assert result.url_type == "email"
        # '/' → '-', '=' stays but is percent-encoded to survive URL parsing.
        assert result.graph_endpoint == "/me/messages/AAMkADJi-zOHQAAA%3D"
        # And crucially: no raw '/' after "/me/messages/" — the id is one
        # opaque path segment.
        prefix = "/me/messages/"
        id_segment = result.graph_endpoint[len(prefix) :]
        assert "/" not in id_segment

    def test_owa_itemid_with_plus_normalized_to_underscore(self):
        """'+' in standard base64 maps to '_' in Outlook's URL-safe form."""
        msg_id_encoded = "AAMk%2BZzOHQAAA%3D"  # '%2B' == '+'
        url = f"https://outlook.office365.com/owa/?ItemID={msg_id_encoded}"
        result = resolve_url(url)
        assert result.url_type == "email"
        assert result.graph_endpoint == "/me/messages/AAMk_ZzOHQAAA%3D"


class TestSharePointPage:
    def test_site_endpoint(self):
        url = "https://contoso.sharepoint.com/sites/MyProject/SitePages/Overview.aspx"
        result = resolve_url(url)
        assert result.url_type == "sharepoint_page"
        assert result.graph_endpoint.startswith("/sites/")

    def test_extra_contains_page_filename(self):
        url = "https://contoso.sharepoint.com/sites/MyProject/SitePages/Overview.aspx"
        result = resolve_url(url)
        assert result.extra["page_filename"] == "Overview.aspx"


class TestOneDriveFile:
    def test_personal_file_endpoint(self):
        url = "https://contoso-my.sharepoint.com/personal/user_example_com/Documents/report.xlsx"
        result = resolve_url(url)
        assert result.url_type == "onedrive_file"
        assert result.graph_endpoint.startswith("/users/")
        assert "report.xlsx" in result.graph_endpoint


class TestOneDriveShareLink:
    def test_shares_api_endpoint(self):
        url = "https://contoso-my.sharepoint.com/:x:/r/personal/user_example_com/_layouts/15/Doc.aspx?sourcedoc=%7Babc%7D"
        result = resolve_url(url)
        assert result.url_type == "onedrive_share_link"
        assert result.graph_endpoint.startswith("/shares/u!")


class TestUnknownUrl:
    def test_unknown_url_raises(self):
        with pytest.raises(UrlParseError):
            resolve_url("https://www.google.com/search?q=hello")


class TestDecodeUpn:
    def test_simple_domain(self):
        assert _decode_upn("alice_smith_example_com") == "alice.smith@example.com"

    def test_multi_part_tld_co_uk(self):
        assert _decode_upn("john_smith_company_co_uk") == "john.smith@company.co.uk"

    def test_multi_part_tld_com_au(self):
        assert _decode_upn("user_name_org_com_au") == "user.name@org.com.au"

    def test_short_encoding_unchanged(self):
        assert _decode_upn("ab") == "ab"

    def test_two_parts_unchanged(self):
        assert _decode_upn("user_com") == "user_com"

    def test_single_name_with_domain(self):
        assert _decode_upn("admin_contoso_com") == "admin@contoso.com"


class TestChatThread:
    def test_meeting_chat_url_parses(self):
        url = (
            "https://teams.microsoft.com/l/chat/"
            "19:meeting_ZmMxNDhi@thread.v2/conversations"
            "?context=%7B%22contextType%22%3A%22chat%22%7D"
        )
        result = resolve_url(url)
        assert result.url_type == "chat_thread"
        assert result.extra["chat_id"] == "19:meeting_ZmMxNDhi@thread.v2"
        assert result.required_scope == "Chat.ReadWrite"

    def test_one_to_one_chat_url_parses(self):
        url = "https://teams.microsoft.com/l/chat/19:abc123@unq.gbl.spaces/conversations"
        result = resolve_url(url)
        assert result.url_type == "chat_thread"
        assert result.extra["chat_id"] == "19:abc123@unq.gbl.spaces"

    def test_group_chat_url_parses(self):
        url = "https://teams.microsoft.com/l/chat/19:groupabc@thread.v2/conversations"
        result = resolve_url(url)
        assert result.url_type == "chat_thread"
        assert result.extra["chat_id"] == "19:groupabc@thread.v2"

    def test_chat_message_url_does_not_match_chat_thread(self):
        # /l/message/ path must continue to match chat_message, not chat_thread
        url = "https://teams.microsoft.com/l/message/19:somechat@unq.gbl.spaces/1234567890.123456"
        result = resolve_url(url)
        assert result.url_type == "chat_message"

    def test_endpoint_template_is_chats_chat_id(self):
        url = "https://teams.microsoft.com/l/chat/19:abc@unq.gbl.spaces/conversations"
        result = resolve_url(url)
        assert result.graph_endpoint == "/chats/19:abc@unq.gbl.spaces"

    def test_malformed_chat_url_raises(self):
        url = "https://teams.microsoft.com/l/chat/not-a-chat-id"
        with pytest.raises(UrlParseError):
            resolve_url(url)

    def test_percent_encoded_chat_url_parses(self):
        # Graph's /me/chats endpoint can return webUrl with %3A (`:`) and %40 (`@`) encoded
        url = "https://teams.microsoft.com/l/chat/19%3Ameeting_ZmMxNDhi%40thread.v2/conversations"
        result = resolve_url(url)
        assert result.url_type == "chat_thread"
        assert result.extra["chat_id"] == "19:meeting_ZmMxNDhi@thread.v2"

    def test_build_chat_thread_url_round_trips_through_resolve(self):
        # build_chat_thread_url is the inverse of the chat_thread parse pattern:
        # a URL it produces must parse back to the same chat_id.
        chat_id = "19:abc123@unq.gbl.spaces"
        built = build_chat_thread_url(chat_id)
        assert resolve_url(built).extra["chat_id"] == chat_id

    def test_build_chat_thread_url_empty_for_falsy_id(self):
        assert build_chat_thread_url("") == ""

    def test_build_chat_thread_url_omits_tenant_id(self):
        # The reconstructed URL intentionally omits ?tenantId= (documented
        # trade-off in build_chat_thread_url). Lock that so it isn't "fixed"
        # by accident without revisiting the extra-Graph-call cost.
        assert "tenantId" not in build_chat_thread_url("19:abc@thread.v2")

    def test_chat_thread_regex_prefix_matches_url_base(self):
        # build and parse hardcode the l/chat/ literal separately (a regex
        # can't interpolate the constant). Guard against drift: the host+path
        # of the base URL must appear in the chat_thread parse pattern. The
        # regex omits the https:// scheme by design (so .search matches
        # anywhere), so compare the post-scheme portion.
        host_path = _CHAT_THREAD_URL_BASE.split("://", 1)[1]
        chat_thread_pattern = next(pat for name, pat, _ in _PATTERNS if name == "chat_thread")
        assert re.escape(host_path) in chat_thread_pattern.pattern
