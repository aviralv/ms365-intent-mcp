"""Validation tests for TranscriptPayload's mutually-exclusive input modes.

Modes: url | name | (item_id+drive_id+site_root) | list=true — exactly one.
"""

from __future__ import annotations

import pytest
from pydantic import ValidationError

from ms365_intent_mcp.intent.transcript.schemas import TranscriptPayload


def test_url_alone_is_valid():
    p = TranscriptPayload(url="https://example/rec")
    assert p.url == "https://example/rec"


def test_name_alone_is_valid():
    assert TranscriptPayload(name="Sprint Review").name == "Sprint Review"


def test_list_alone_is_valid():
    assert TranscriptPayload(list=True).list is True


def test_full_coords_triple_is_valid():
    p = TranscriptPayload(item_id="i", drive_id="d", site_root="s")
    assert (p.item_id, p.drive_id, p.site_root) == ("i", "d", "s")


def test_no_input_is_rejected():
    with pytest.raises(ValidationError):
        TranscriptPayload()


def test_two_modes_are_mutually_exclusive():
    with pytest.raises(ValidationError):
        TranscriptPayload(url="u", name="n")


def test_list_plus_name_is_rejected():
    with pytest.raises(ValidationError):
        TranscriptPayload(name="n", list=True)


def test_partial_coords_triple_is_rejected():
    """item_id without drive_id/site_root can't self-locate for Vroom."""
    with pytest.raises(ValidationError):
        TranscriptPayload(item_id="i")


def test_coords_plus_name_is_rejected():
    with pytest.raises(ValidationError):
        TranscriptPayload(item_id="i", drive_id="d", site_root="s", name="n")
