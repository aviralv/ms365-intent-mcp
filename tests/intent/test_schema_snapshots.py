"""Golden JSON Schema snapshots for compose_v1 payloads and response models.

8 snapshot tests — one per schema. First run writes the JSON file;
subsequent runs assert byte-equality.  Use ``SNAPSHOT_UPDATE=1`` to
refresh a snapshot after an intentional schema change.

Gate confirmation: ``compose_payload.json`` must contain both ``oneOf``
and ``discriminator`` to confirm that FastMCP/Pydantic renders the
discriminated union correctly (Task 1 gate, re-verified here).
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest
from pydantic import TypeAdapter

from ms365_intent_mcp.intent._shared import ErrorResponse
from ms365_intent_mcp.intent.compose.schemas import (
    ComposeEmail,
    ComposeEvent,
    ComposePayload,
    ComposeTeamsMessage,
    EmailDraftCreated,
    EventCreated,
    TeamsMessageSent,
)

_SNAPSHOT_DIR = Path(__file__).parent / "snapshots" / "schemas"


# ---------------------------------------------------------------------------
# Payload union
# ---------------------------------------------------------------------------


def test_compose_payload_snapshot(snapshot: Any) -> None:
    schema = TypeAdapter(ComposePayload).json_schema()
    snapshot("compose_payload", schema)

    # Gate: discriminated union must render with both keywords
    on_disk = json.loads((_SNAPSHOT_DIR / "compose_payload.json").read_text())
    assert "oneOf" in on_disk, "compose_payload.json must contain 'oneOf'"
    assert "discriminator" in on_disk, "compose_payload.json must contain 'discriminator'"


# ---------------------------------------------------------------------------
# Payload variants
# ---------------------------------------------------------------------------


def test_compose_email_variant(snapshot: Any) -> None:
    snapshot("compose_email", ComposeEmail.model_json_schema())


def test_compose_event_variant(snapshot: Any) -> None:
    snapshot("compose_event", ComposeEvent.model_json_schema())


def test_compose_teams_message_variant(snapshot: Any) -> None:
    snapshot("compose_teams_message", ComposeTeamsMessage.model_json_schema())


# ---------------------------------------------------------------------------
# Response models
# ---------------------------------------------------------------------------


def test_email_draft_created(snapshot: Any) -> None:
    snapshot("email_draft_created", EmailDraftCreated.model_json_schema())


def test_event_created(snapshot: Any) -> None:
    snapshot("event_created", EventCreated.model_json_schema())


def test_teams_message_sent(snapshot: Any) -> None:
    snapshot("teams_message_sent", TeamsMessageSent.model_json_schema())


def test_error_response(snapshot: Any) -> None:
    snapshot("error_response", ErrorResponse.model_json_schema())
