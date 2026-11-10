"""Synthetic calendar updates only: no credentials or live Graph writes."""

from contextlib import asynccontextmanager
from copy import deepcopy
from datetime import UTC, datetime
from unittest.mock import AsyncMock, MagicMock

import pytest
from fastmcp import Client, FastMCP
from jsonschema import validate
from pydantic import ValidationError

from ms365_intent_mcp.composers.compose import ComposeType, compose_action
from ms365_intent_mcp.graph import GraphAPIError, GraphClient
from ms365_intent_mcp.intent._helpers import idempotency_clear
from ms365_intent_mcp.intent._shared import ErrorResponse
from ms365_intent_mcp.intent.compose.impl import _compose_impl
from ms365_intent_mcp.intent.compose.register import register
from ms365_intent_mcp.intent.compose.schemas import ComposeEvent, EventUpdated
from ms365_intent_mcp.permissions import PermissionRegistry


@pytest.fixture
def deps():
    idempotency_clear()
    event = {
        "id": "synthetic/event+id=",
        "subject": "Synthetic Focus Time Blocker",
        "start": {"dateTime": "2026-11-13T12:00:00.0000000", "timeZone": "UTC"},
        "end": {"dateTime": "2026-11-13T15:00:00.0000000", "timeZone": "UTC"},
        "body": {"contentType": "HTML", "content": "Preserve Teams join metadata"},
        "attendees": [{"emailAddress": {"address": "synthetic@example.com"}}],
        "location": {"displayName": "Desk", "uniqueId": "synthetic-location"},
        "recurrence": {"pattern": {"type": "weekly"}},
        "type": "occurrence",
        "isOnlineMeeting": True,
        "onlineMeeting": {"joinUrl": "https://teams.microsoft.com/synthetic"},
        "webLink": "https://outlook.office.com/calendar/item/synthetic",
        "@odata.etag": 'W/"synthetic-version"',
    }
    graph = AsyncMock(spec=GraphClient)
    graph.get.return_value = event

    async def patch_event(endpoint, patch, headers=None):
        # Model Graph PATCH's omitted-field preservation and UTC response.
        result = deepcopy(event)
        result.update(patch)
        return result

    graph.patch.side_effect = patch_event
    permissions = PermissionRegistry(["Calendars.ReadWrite"])
    ctx = MagicMock()
    context = {
        "config": MagicMock(default_timezone="America/Los_Angeles"),
        "client": graph,
        "permissions": permissions,
    }
    ctx.request_context.lifespan_context = context
    yield ctx, graph, permissions, context
    idempotency_clear()


def payload(**changes):
    return ComposeEvent.model_validate(
        {"type": "event", "mode": "update", "event_id": "synthetic/event+id=", **changes}
    )


@pytest.mark.asyncio
async def test_resize_targets_exact_id_and_preserves_omitted_details(deps):
    ctx, graph, _, _ = deps
    original = deepcopy(graph.get.return_value)
    response = await _compose_impl(
        ctx, payload(start="2026-11-13T14:00:00", timezone="Europe/Berlin")
    )
    assert isinstance(response, EventUpdated)
    endpoint = "/me/events/synthetic%2Fevent%2Bid%3D"
    graph.get.assert_awaited_once_with(endpoint, headers=GraphClient.calendar_headers("UTC"))
    graph.patch.assert_awaited_once_with(
        endpoint,
        {"start": {"dateTime": "2026-11-13T13:00:00", "timeZone": "UTC"}},
        headers={"Prefer": 'outlook.timezone="UTC"', "If-Match": original["@odata.etag"]},
    )
    graph.post.assert_not_awaited()
    assert graph.get.return_value == original
    assert response.start.isoformat() == "2026-11-13T14:00:00+01:00"
    assert response.end.isoformat() == "2026-11-13T16:00:00+01:00"
    assert response.timezone == "Europe/Berlin"
    assert "Event updated" in response.rendered_markdown
    assert "14:00:00+01:00 Europe/Berlin" in response.rendered_markdown
    assert "16:00:00+01:00 Europe/Berlin" in response.rendered_markdown
    assert "Calendar:" in response.rendered_markdown
    assert original["id"] not in response.rendered_markdown


@pytest.mark.parametrize(
    "changes",
    [
        {},
        {"event_id": None, "subject": "New"},
        {"event_id": " ", "subject": "New"},
        {"subject": None},
        {"subject": " "},
        {"start": None},
        {"location": None},
        {"timezone": "Europe/Berlin"},
        {"start": "2026-11-13T14:00:00"},
        {"start": "2026-11-13T14:00:00", "timezone": "CET-not-IANA"},
        {"start": "2026-11-13T14:00:00Z", "timezone": None},
        {"subject": "New", "attendees": []},
        {"subject": "New", "is_online_meeting": False},
        {"subject": "New", "body": "replacement"},
        {"subject": "New", "to": []},
        {"subject": "New", "comment": "ignored"},
    ],
)
def test_update_schema_rejects_unsafe_or_empty_inputs(changes):
    with pytest.raises(ValidationError):
        payload(**changes)


def test_missing_event_id_is_rejected():
    with pytest.raises(ValidationError, match="requires 'event_id'"):
        ComposeEvent.model_validate({"type": "event", "mode": "update", "subject": "New"})


@pytest.mark.asyncio
@pytest.mark.parametrize(
    "changes,message",
    [
        ({"start": "2026-11-13T16:00:00+01:00"}, "end must be after start"),
        ({"end": "2026-11-13T12:00:00+01:00"}, "end must be after start"),
        (
            {"start": "2026-11-13T17:00:00+01:00", "end": "2026-11-13T16:00:00Z"},
            "end must be after start",
        ),
        ({"start": "2026-03-29T02:30:00", "timezone": "Europe/Berlin"}, "Nonexistent"),
        ({"start": "2026-10-25T02:30:00", "timezone": "Europe/Berlin"}, "Ambiguous"),
        ({"start": "2026-11-13T13:00:00+01:00"}, "No event fields would change"),
        ({"subject": "Synthetic Focus Time Blocker"}, "No event fields would change"),
    ],
)
async def test_runtime_validation_blocks_patch(deps, changes, message):
    ctx, graph, _, _ = deps
    response = await _compose_impl(ctx, payload(**changes))
    assert isinstance(response, ErrorResponse)
    assert response.code == "validation_error"
    assert message in response.message
    graph.patch.assert_not_awaited()


@pytest.mark.asyncio
@pytest.mark.parametrize(
    "start,end,timezone,expected_start,expected_end",
    [
        (
            "2026-07-10T14:00:00",
            "2026-07-10T16:00:00",
            "Europe/Berlin",
            "2026-07-10T12:00:00",
            "2026-07-10T14:00:00",
        ),
        (
            "2026-03-29T01:30:00",
            "2026-03-29T03:30:00",
            "Europe/Berlin",
            "2026-03-29T00:30:00",
            "2026-03-29T01:30:00",
        ),
        (
            "2026-10-25T02:30:00+02:00",
            "2026-10-25T02:30:00+01:00",
            "Europe/Berlin",
            "2026-10-25T00:30:00",
            "2026-10-25T01:30:00",
        ),
        (
            "2026-11-13T14:00:00Z",
            "2026-11-13T16:00:00+01:00",
            "Europe/Berlin",
            "2026-11-13T14:00:00",
            "2026-11-13T15:00:00",
        ),
        (
            "2026-11-13T14:00:00+01:00",
            "2026-11-13T17:00:00+01:00",
            None,
            "2026-11-13T13:00:00",
            "2026-11-13T16:00:00",
        ),
    ],
)
async def test_reschedule_normalizes_instants_without_dst_or_config_shifts(
    deps, start, end, timezone, expected_start, expected_end
):
    ctx, graph, _, _ = deps
    changes = {"start": start, "end": end}
    if timezone:
        changes["timezone"] = timezone
    response = await _compose_impl(ctx, payload(**changes))
    assert isinstance(response, EventUpdated)
    patch = graph.patch.call_args.args[1]
    assert patch["start"] == {"dateTime": expected_start, "timeZone": "UTC"}
    # An unchanged end is omitted, even if supplied with a different offset.
    if expected_end == "2026-11-13T15:00:00":
        assert "end" not in patch
    else:
        assert patch["end"] == {"dateTime": expected_end, "timeZone": "UTC"}
    assert response.start.astimezone(UTC) == datetime.fromisoformat(expected_start).replace(
        tzinfo=UTC
    )
    assert response.end.astimezone(UTC) == datetime.fromisoformat(expected_end).replace(tzinfo=UTC)
    if start == "2026-10-25T02:30:00+02:00":
        assert "02:30:00+02:00 Europe/Berlin" in response.rendered_markdown
        assert "02:30:00+01:00 Europe/Berlin" in response.rendered_markdown


@pytest.mark.asyncio
async def test_end_only_and_explicit_location_clear_do_not_send_defaults(deps):
    ctx, graph, _, _ = deps
    response = await _compose_impl(ctx, payload(end="2026-11-13T17:00:00+01:00", location=""))
    assert isinstance(response, EventUpdated)
    assert graph.patch.call_args.args[1] == {
        "end": {"dateTime": "2026-11-13T16:00:00", "timeZone": "UTC"},
        "location": {"displayName": ""},
    }
    assert response.start.isoformat() == "2026-11-13T12:00:00+00:00"


@pytest.mark.asyncio
async def test_subject_update_preserves_times_without_twelve_hour_create_limit(deps):
    ctx, graph, _, _ = deps
    graph.get.return_value["end"]["dateTime"] = "2026-11-14T15:00:00"
    graph.get.return_value.pop("@odata.etag")
    response = await _compose_impl(ctx, payload(subject="Renamed"))
    assert isinstance(response, EventUpdated)
    assert graph.patch.call_args.args[1] == {"subject": "Renamed"}
    assert "If-Match" not in graph.patch.call_args.kwargs["headers"]


@pytest.mark.asyncio
@pytest.mark.parametrize("field,value", [("type", "seriesMaster"), ("isAllDay", True)])
async def test_rejects_series_master_and_all_day_time_updates(deps, field, value):
    ctx, graph, _, _ = deps
    graph.get.return_value[field] = value
    response = await _compose_impl(ctx, payload(start="2026-11-13T14:00:00+01:00"))
    assert response.code == "validation_error"
    graph.patch.assert_not_awaited()


@pytest.mark.asyncio
@pytest.mark.parametrize("operation", ["get", "patch"])
@pytest.mark.parametrize("status", [403, 404, 412, 429, 503])
async def test_graph_failures_are_errors_not_success_or_cached(deps, operation, status):
    ctx, graph, _, _ = deps
    getattr(graph, operation).side_effect = GraphAPIError(status, "SyntheticError", "failure")
    request = payload(subject="Renamed", idempotency_key="synthetic-retry")
    for _ in range(2):
        response = await _compose_impl(ctx, request)
        assert isinstance(response, ErrorResponse)
        assert response.code == ("rate_limited" if status in (429, 503) else "graph_api_error")
        assert response.retryable is (status in (429, 503))
    assert getattr(graph, operation).await_count == 2
    if operation == "get":
        graph.patch.assert_not_awaited()


@pytest.mark.asyncio
async def test_missing_scope_never_reads_or_mutates(deps):
    ctx, graph, _, context = deps
    context["permissions"] = PermissionRegistry([])
    response = await _compose_impl(ctx, payload(subject="Renamed"))
    assert isinstance(response, ErrorResponse)
    assert response.code == "insufficient_scope"
    graph.get.assert_not_awaited()
    graph.patch.assert_not_awaited()


@pytest.mark.asyncio
async def test_bad_current_timing_and_empty_patch_response_fail_closed(deps):
    ctx, graph, _, _ = deps
    graph.get.return_value["start"]["timeZone"] = "unknown"
    response = await _compose_impl(ctx, payload(subject="Renamed"))
    assert response.code == "graph_api_error"
    graph.patch.assert_not_awaited()
    graph.get.return_value["start"]["timeZone"] = "UTC"
    graph.patch.side_effect = None
    graph.patch.return_value = {}
    response = await _compose_impl(ctx, payload(subject="Renamed"))
    assert response.code == "graph_api_error"


@pytest.mark.asyncio
async def test_idempotent_update_retries_are_cached(deps):
    ctx, graph, _, _ = deps
    request = payload(subject="Renamed", idempotency_key="synthetic-success")
    first = await _compose_impl(ctx, request)
    second = await _compose_impl(ctx, request)
    assert first == second
    graph.patch.assert_awaited_once()


@pytest.mark.asyncio
async def test_dispatch_rejects_missing_id_without_graph_access(deps):
    _, graph, permissions, _ = deps
    from ms365_intent_mcp.intent._helpers import IntentError

    with pytest.raises(IntentError, match="exact event_id"):
        await compose_action(graph, permissions, ComposeType.EVENT_UPDATE, {"start": "ignored"})
    graph.get.assert_not_awaited()
    graph.patch.assert_not_awaited()


@pytest.mark.asyncio
@pytest.mark.parametrize("mode", ["update", "create", "forward", "error"])
async def test_actual_mcp_schema_accepts_new_mode_errors_and_existing_modes(deps, mode):
    _, graph, _, context = deps

    @asynccontextmanager
    async def lifespan(server):
        yield context

    server = FastMCP("synthetic-compose-test", lifespan=lifespan)
    register(server)
    if mode in ("update", "error"):
        request = payload(subject="Renamed").model_dump(mode="json", exclude_unset=True)
        expected = "event_updated" if mode == "update" else "error"
        if mode == "error":
            graph.patch.side_effect = GraphAPIError(412, "PreconditionFailed", "changed")
    elif mode == "create":
        request = {
            "type": "event",
            "subject": "Synthetic create",
            "start": "2026-11-13T13:00:00",
            "end": "2026-11-13T14:00:00",
            "timezone": "Europe/Berlin",
        }
        graph.post.return_value = {**graph.get.return_value, "subject": "Synthetic create"}
        expected = "event_created"
    else:
        request = {
            "type": "event",
            "mode": "forward",
            "event_id": "synthetic-event",
            "to": [{"email": "synthetic@example.com"}],
        }
        graph.post.return_value = {}
        expected = "event_forwarded"
    async with Client(server) as client:
        tool = next(t for t in await client.list_tools() if t.name == "compose")
        result = await client.call_tool("compose", {"payload": request})
    validate(
        instance=result.structured_content, schema=tool.model_dump(by_alias=True)["outputSchema"]
    )
    # Preserve compose's pre-existing inferred union wrapper, unlike meeting's flat contract.
    assert result.structured_content["result"]["type"] == expected
    if mode in ("create", "forward"):
        graph.patch.assert_not_awaited()
        graph.get.assert_not_awaited()
