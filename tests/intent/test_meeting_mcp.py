"""Exercise meeting's advertised response contract over an actual MCP connection."""

from contextlib import asynccontextmanager
from unittest.mock import AsyncMock, MagicMock

import pytest
from fastmcp import Client, FastMCP
from jsonschema import validate

from ms365_intent_mcp.graph import GraphAPIError, GraphClient
from ms365_intent_mcp.intent.meeting.register import register
from ms365_intent_mcp.permissions import PermissionRegistry


@pytest.fixture
def meeting_server():
    event = {
        "id": "synthetic-occurrence-id",
        "type": "occurrence",
        "subject": "Synthetic Weekly Sync",
        "start": {"dateTime": "2026-11-09T09:00:00", "timeZone": "UTC"},
        "end": {"dateTime": "2026-11-09T09:30:00", "timeZone": "UTC"},
        "organizer": {"emailAddress": {"name": "Synthetic Organizer"}},
        "attendees": [],
    }
    graph = AsyncMock(spec=GraphClient)
    graph.get.return_value = event
    graph.get_all.return_value = ([event], False)
    deps = {
        "config": MagicMock(default_timezone="UTC"),
        "client": graph,
        "permissions": PermissionRegistry(["Calendars.ReadWrite"]),
    }

    @asynccontextmanager
    async def lifespan(server):
        yield deps

    server = FastMCP("synthetic-meeting-test", lifespan=lifespan)
    register(server)
    return server, graph


@pytest.mark.asyncio
@pytest.mark.parametrize(
    "identifier,requested_date,expected_type",
    [
        ("Synthetic Weekly Sync", "2026-11-09", "meeting_detail"),
        ("Synthetic Weekly Sync", "2026-11-10", "error"),
        ("synthetic-occurrence-id", "2026-11-09", "meeting_detail"),
        ("synthetic-occurrence-id", "2026-11-10", "error"),
        ("synthetic-occurrence-id", None, "meeting_detail"),
    ],
)
async def test_success_and_date_misses_match_advertised_output_schema(
    meeting_server, identifier, requested_date, expected_type
):
    server, _ = meeting_server
    payload = {"identifier": identifier}
    if requested_date is not None:
        payload["date"] = requested_date

    async with Client(server) as client:
        tool = next(t for t in await client.list_tools() if t.name == "meeting")
        result = await client.call_tool("meeting", {"payload": payload})

    data = result.structured_content
    validate(instance=data, schema=tool.model_dump(by_alias=True)["outputSchema"])
    assert "result" not in data, "Successful responses must not gain a FastMCP union wrapper"
    assert data["type"] == expected_type
    if expected_type == "error":
        assert data["code"] == "not_found"
        assert data["retryable"] is False
        assert requested_date in data["message"]
        assert "id" not in data
        assert "start" not in data
    else:
        assert data["id"] == "synthetic-occurrence-id"
        assert data["start"].startswith("2026-11-09")
        assert "rendered_markdown" in data


@pytest.mark.asyncio
@pytest.mark.parametrize("status", [403, 429])
async def test_graph_errors_also_match_advertised_output_schema(meeting_server, status):
    server, graph = meeting_server
    graph.get_all.side_effect = GraphAPIError(status, "SyntheticError", "Synthetic failure")
    async with Client(server) as client:
        tool = next(t for t in await client.list_tools() if t.name == "meeting")
        result = await client.call_tool(
            "meeting", {"payload": {"identifier": "Synthetic Weekly Sync", "date": "2026-11-09"}}
        )
    data = result.structured_content
    validate(instance=data, schema=tool.model_dump(by_alias=True)["outputSchema"])
    assert "result" not in data
    assert data["type"] == "error"
    assert data["code"] == ("rate_limited" if status == 429 else "graph_api_error")
    assert data["retryable"] is (status == 429)


@pytest.mark.asyncio
async def test_series_master_validation_error_matches_output_schema(meeting_server):
    server, graph = meeting_server
    graph.get.return_value = {**graph.get.return_value, "type": "seriesMaster"}
    async with Client(server) as client:
        tool = next(t for t in await client.list_tools() if t.name == "meeting")
        result = await client.call_tool(
            "meeting", {"payload": {"identifier": "synthetic-occurrence-id", "date": "2026-11-09"}}
        )
    data = result.structured_content
    validate(instance=data, schema=tool.model_dump(by_alias=True)["outputSchema"])
    assert data["type"] == "error"
    assert data["code"] == "validation_error"
    assert "result" not in data


@pytest.mark.asyncio
@pytest.mark.parametrize("value", ["20261109", "2026-02-30"])
async def test_malformed_dates_are_input_errors_without_graph_access(meeting_server, value):
    server, graph = meeting_server
    async with Client(server) as client:
        result = await client.call_tool(
            "meeting",
            {"payload": {"identifier": "Synthetic Weekly Sync", "date": value}},
            raise_on_error=False,
        )
    assert result.is_error is True
    assert result.structured_content is None
    graph.get.assert_not_awaited()
    graph.get_all.assert_not_awaited()
