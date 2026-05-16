"""Tests for URL resolver dispatch table."""

import pytest

from ms365_intent_mcp.resolver import resolve_url, UrlParseError, ResolvedUrl


class TestChannelMessageUrls:
    def test_teams_channel_message(self):
        url = "https://teams.microsoft.com/l/message/19:abc123@thread.tacv2/1234567890.123456?context=%7B%22tid%22%3A%22tenant%22%7D"
        result = resolve_url(url)
        assert result.url_type == "channel_message"
        assert "19:abc123@thread.tacv2" in result.graph_endpoint

    def test_teams_chat_message(self):
        url = "https://teams.microsoft.com/l/message/19:abc123def@unq.gbl.spaces/1234567890.123456"
        result = resolve_url(url)
        assert result.url_type == "chat_message"

    def test_teams_meeting_link(self):
        url = "https://teams.microsoft.com/l/meetup-join/19:meeting_abc123@thread.v2/0?context=%7B%22Tid%22%3A%22tenant%22%7D"
        result = resolve_url(url)
        assert result.url_type == "meeting"

    def test_outlook_mail_deeplink(self):
        url = "https://outlook.office365.com/mail/id/AAkALgAAAAAAHYQ..."
        result = resolve_url(url)
        assert result.url_type == "email"

    def test_sharepoint_page(self):
        url = "https://sap.sharepoint.com/sites/MyProject/SitePages/Overview.aspx"
        result = resolve_url(url)
        assert result.url_type == "sharepoint_page"

    def test_onedrive_file(self):
        url = "https://sap-my.sharepoint.com/personal/user_sap_com/Documents/report.xlsx"
        result = resolve_url(url)
        assert result.url_type == "onedrive_file"

    def test_onedrive_share_link(self):
        url = "https://sap-my.sharepoint.com/:x:/r/personal/user_sap_com/_layouts/15/Doc.aspx?sourcedoc=%7Babc%7D"
        result = resolve_url(url)
        assert result.url_type == "onedrive_share_link"


class TestUnknownUrl:
    def test_unknown_url_raises(self):
        with pytest.raises(UrlParseError):
            resolve_url("https://www.google.com/search?q=hello")
