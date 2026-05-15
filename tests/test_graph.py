"""Tests for GraphClient response handling."""

import pytest
from unittest.mock import AsyncMock, patch

from ms365_intent_mcp.graph import GraphAPIError
from tests.conftest import make_graph_client, make_graph_response


class TestHandleResponse:
    def test_200_returns_json(self):
        response = make_graph_response(200, json_body={"value": [1, 2]})
        result = make_graph_client()._handle_response(response)
        assert result == {"value": [1, 2]}

    def test_204_returns_empty_dict(self):
        response = make_graph_response(204)
        result = make_graph_client()._handle_response(response)
        assert result == {}

    def test_401_raises_auth_error(self):
        response = make_graph_response(
            401, json_body={"error": {"code": "InvalidToken", "message": "expired"}}
        )
        with pytest.raises(GraphAPIError) as exc_info:
            make_graph_client()._handle_response(response)
        assert exc_info.value.status_code == 401

    def test_404_raises_not_found(self):
        response = make_graph_response(
            404, json_body={"error": {"code": "NotFound", "message": "gone"}}
        )
        with pytest.raises(GraphAPIError) as exc_info:
            make_graph_client()._handle_response(response)
        assert exc_info.value.status_code == 404

    def test_429_includes_status(self):
        response = make_graph_response(
            429, json_body={"error": {"code": "TooManyRequests", "message": "slow down"}}
        )
        with pytest.raises(GraphAPIError) as exc_info:
            make_graph_client()._handle_response(response)
        assert exc_info.value.status_code == 429


class TestGraphAPIError:
    def test_attributes(self):
        err = GraphAPIError(500, "InternalError", "broke")
        assert err.status_code == 500
        assert err.error_code == "InternalError"
        assert "500" in str(err)


class TestGetAll:
    @pytest.mark.asyncio
    async def test_single_page_no_next_link(self):
        client = make_graph_client()
        async with client:
            with patch.object(client, "get", new_callable=AsyncMock) as mock_get:
                mock_get.return_value = {"value": [{"id": "1"}, {"id": "2"}]}
                items, has_more = await client.get_all("/me/messages")
        assert items == [{"id": "1"}, {"id": "2"}]
        assert has_more is False
        mock_get.assert_called_once_with("/me/messages", params=None, headers=None)

    @pytest.mark.asyncio
    async def test_follows_next_link(self):
        client = make_graph_client()
        async with client:
            with patch.object(client, "get", new_callable=AsyncMock) as mock_get:
                mock_get.side_effect = [
                    {
                        "value": [{"id": "1"}],
                        "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/messages?$skip=1",
                    },
                    {"value": [{"id": "2"}]},
                ]
                items, has_more = await client.get_all("/me/messages")
        assert items == [{"id": "1"}, {"id": "2"}]
        assert has_more is False
        assert mock_get.call_count == 2

    @pytest.mark.asyncio
    async def test_respects_max_pages_and_signals_has_more(self):
        client = make_graph_client()
        async with client:
            with patch.object(client, "get", new_callable=AsyncMock) as mock_get:
                mock_get.return_value = {
                    "value": [{"id": "x"}],
                    "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/messages?$skip=1",
                }
                items, has_more = await client.get_all("/me/messages", max_pages=2)
        assert len(items) == 2
        assert has_more is True
        assert mock_get.call_count == 2

    @pytest.mark.asyncio
    async def test_empty_value_list(self):
        client = make_graph_client()
        async with client:
            with patch.object(client, "get", new_callable=AsyncMock) as mock_get:
                mock_get.return_value = {"value": []}
                items, has_more = await client.get_all("/me/messages")
        assert items == []
        assert has_more is False
