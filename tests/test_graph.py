"""Tests for GraphClient response handling."""

from unittest.mock import AsyncMock, patch

import httpx
import pytest

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



class TestGetContent:
    @pytest.mark.asyncio
    async def test_returns_bytes_on_200(self):
        client = make_graph_client()
        async with client:
            fake_response = httpx.Response(
                200,
                content=b"hello file content",
                headers={"content-type": "text/plain"},
                request=httpx.Request("GET", "https://graph.microsoft.com/v1.0/test"),
            )
            with patch.object(client._client, "get", new_callable=AsyncMock) as mock_get:
                mock_get.return_value = fake_response
                result = await client.get_content("/me/drive/items/123/content")
        assert result == b"hello file content"

    @pytest.mark.asyncio
    async def test_follows_allowed_redirect(self):
        client = make_graph_client()
        async with client:
            redirect_response = httpx.Response(
                302,
                headers={"location": "https://files.sharepoint.com/download/file.xlsx"},
                request=httpx.Request("GET", "https://graph.microsoft.com/v1.0/test"),
            )
            final_response = httpx.Response(
                200,
                content=b"file data",
                headers={"content-type": "application/octet-stream"},
                request=httpx.Request("GET", "https://files.sharepoint.com/download/file.xlsx"),
            )
            with patch.object(client._client, "get", new_callable=AsyncMock) as mock_get:
                mock_get.side_effect = [redirect_response, final_response]
                result = await client.get_content("/me/drive/items/123/content")
        assert result == b"file data"
        assert mock_get.call_count == 2

    @pytest.mark.asyncio
    async def test_blocks_disallowed_redirect(self):
        client = make_graph_client()
        async with client:
            redirect_response = httpx.Response(
                302,
                headers={"location": "https://evil.com/steal"},
                request=httpx.Request("GET", "https://graph.microsoft.com/v1.0/test"),
            )
            with patch.object(client._client, "get", new_callable=AsyncMock) as mock_get:
                mock_get.return_value = redirect_response
                with pytest.raises(GraphAPIError) as exc_info:
                    await client.get_content("/me/drive/items/123/content")
        assert exc_info.value.status_code == 403
        assert "RedirectBlocked" in exc_info.value.error_code

    @pytest.mark.asyncio
    async def test_raises_on_404(self):
        client = make_graph_client()
        async with client:
            fake_response = httpx.Response(
                404,
                content=b'{"error":{"code":"NotFound","message":"gone"}}',
                headers={"content-type": "application/json"},
                request=httpx.Request("GET", "https://graph.microsoft.com/v1.0/test"),
            )
            with patch.object(client._client, "get", new_callable=AsyncMock) as mock_get:
                mock_get.return_value = fake_response
                with pytest.raises(GraphAPIError) as exc_info:
                    await client.get_content("/me/drive/items/missing/content")
        assert exc_info.value.status_code == 404


class TestRetryAfter:
    @pytest.mark.asyncio
    async def test_retries_on_429_with_retry_after_header(self):
        client = make_graph_client()
        async with client:
            throttled_response = httpx.Response(
                429,
                content=b'{"error":{"code":"TooManyRequests","message":"slow down"}}',
                headers={"content-type": "application/json", "Retry-After": "1"},
                request=httpx.Request("GET", "https://graph.microsoft.com/v1.0/me"),
            )
            success_response = httpx.Response(
                200,
                content=b'{"value": []}',
                headers={"content-type": "application/json"},
                request=httpx.Request("GET", "https://graph.microsoft.com/v1.0/me"),
            )
            with patch.object(client._client, "get", new_callable=AsyncMock) as mock_get:
                mock_get.side_effect = [throttled_response, success_response]
                with patch(
                    "ms365_intent_mcp.graph.asyncio.sleep", new_callable=AsyncMock
                ) as mock_sleep:
                    result = await client.get("/me")
        assert result == {"value": []}
        assert mock_get.call_count == 2
        mock_sleep.assert_called_once_with(1)

    @pytest.mark.asyncio
    async def test_retries_on_503(self):
        client = make_graph_client()
        async with client:
            error_response = httpx.Response(
                503,
                content=b'{"error":{"code":"ServiceUnavailable","message":"try again"}}',
                headers={"content-type": "application/json"},
                request=httpx.Request("GET", "https://graph.microsoft.com/v1.0/me"),
            )
            success_response = httpx.Response(
                200,
                content=b'{"id": "123"}',
                headers={"content-type": "application/json"},
                request=httpx.Request("GET", "https://graph.microsoft.com/v1.0/me"),
            )
            with patch.object(client._client, "get", new_callable=AsyncMock) as mock_get:
                mock_get.side_effect = [error_response, success_response]
                with patch(
                    "ms365_intent_mcp.graph.asyncio.sleep", new_callable=AsyncMock
                ) as mock_sleep:
                    result = await client.get("/me")
        assert result == {"id": "123"}
        mock_sleep.assert_called_once_with(1)

    @pytest.mark.asyncio
    async def test_raises_if_retry_also_fails(self):
        client = make_graph_client()
        async with client:
            throttled = httpx.Response(
                429,
                content=b'{"error":{"code":"TooManyRequests","message":"slow down"}}',
                headers={"content-type": "application/json", "Retry-After": "2"},
                request=httpx.Request("GET", "https://graph.microsoft.com/v1.0/me"),
            )
            with patch.object(client._client, "get", new_callable=AsyncMock) as mock_get:
                mock_get.return_value = throttled
                with patch(
                    "ms365_intent_mcp.graph.asyncio.sleep", new_callable=AsyncMock
                ) as mock_sleep:
                    with pytest.raises(GraphAPIError) as exc_info:
                        await client.get("/me")
        assert exc_info.value.status_code == 429
        assert mock_get.call_count == 2
        mock_sleep.assert_called_once_with(2)

    @pytest.mark.asyncio
    async def test_caps_retry_after_at_10(self):
        client = make_graph_client()
        async with client:
            throttled = httpx.Response(
                429,
                content=b'{"error":{"code":"TooManyRequests","message":"slow"}}',
                headers={"content-type": "application/json", "Retry-After": "60"},
                request=httpx.Request("GET", "https://graph.microsoft.com/v1.0/me"),
            )
            success = httpx.Response(
                200,
                content=b'{"ok": true}',
                headers={"content-type": "application/json"},
                request=httpx.Request("GET", "https://graph.microsoft.com/v1.0/me"),
            )
            with patch.object(client._client, "get", new_callable=AsyncMock) as mock_get:
                mock_get.side_effect = [throttled, success]
                with patch(
                    "ms365_intent_mcp.graph.asyncio.sleep", new_callable=AsyncMock
                ) as mock_sleep:
                    await client.get("/me")
        mock_sleep.assert_called_once_with(10)


class TestGraphClientAbsoluteURL:
    @pytest.mark.asyncio
    async def test_accepts_absolute_graph_url(self):
        from tests.conftest import make_graph_client, make_graph_response

        client = make_graph_client()
        async with client:
            with patch.object(
                client._client,
                "get",
                AsyncMock(return_value=make_graph_response(200, {"value": []})),
            ) as mock_get:
                await client.get("https://graph.microsoft.com/v1.0/me/messages")
            called_url = mock_get.call_args[0][0]
            assert called_url == "https://graph.microsoft.com/v1.0/me/messages"

    @pytest.mark.asyncio
    async def test_rejects_non_graph_absolute_url(self):
        from tests.conftest import make_graph_client

        client = make_graph_client()
        async with client:
            with pytest.raises(ValueError, match="non-Graph host"):
                await client.get("https://evil.example.com/v1.0/me/messages")

    @pytest.mark.asyncio
    async def test_rejects_spoofed_subdomain(self):
        from tests.conftest import make_graph_client

        client = make_graph_client()
        async with client:
            with pytest.raises(ValueError, match="non-Graph host"):
                await client.get("https://graph.microsoft.com.evil.com/v1.0/me/messages")

    @pytest.mark.asyncio
    async def test_relative_path_unchanged(self):
        from tests.conftest import make_graph_client, make_graph_response

        client = make_graph_client()
        async with client:
            with patch.object(
                client._client,
                "get",
                AsyncMock(return_value=make_graph_response(200, {"value": []})),
            ) as mock_get:
                await client.get("/me/messages")
            called_url = mock_get.call_args[0][0]
            assert called_url == "https://graph.microsoft.com/v1.0/me/messages"
