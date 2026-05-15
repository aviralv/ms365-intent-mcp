"""Shared test factories and helpers."""

import json

import httpx

from ms365_intent_mcp.graph import GraphClient, GraphAPIError


def make_graph_response(
    status_code: int,
    json_body: dict | None = None,
    text: str = "",
) -> httpx.Response:
    if json_body is not None:
        raw = json.dumps(json_body).encode()
        headers = {"content-type": "application/json"}
    else:
        raw = text.encode()
        headers = {"content-type": "text/plain"}
    return httpx.Response(
        status_code,
        content=raw,
        headers=headers,
        request=httpx.Request("GET", "https://graph.microsoft.com/v1.0/test"),
    )


def make_graph_client() -> GraphClient:
    return GraphClient("https://graph.microsoft.com/v1.0", lambda: "fake-token")
