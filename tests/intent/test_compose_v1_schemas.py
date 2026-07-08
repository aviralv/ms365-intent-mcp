"""Schema-level tests for compose_v1 discriminated union."""

import pytest
from pydantic import ValidationError

from ms365_intent_mcp.intent.compose.schemas import (
    ComposeEmail,
    ComposeEvent,
    ComposeTeamsMessage,
)


class TestComposeEmail:
    def test_new_email_valid(self):
        m = ComposeEmail.model_validate({
            "type": "email",
            "mode": "new",
            "to": [{"email": "a@b.com"}],
            "subject": "Hi",
            "body": "Hello",
        })
        assert m.mode == "new"

    def test_new_email_missing_to_raises(self):
        with pytest.raises(ValidationError, match="requires both 'to' and 'subject'"):
            ComposeEmail.model_validate({
                "type": "email",
                "mode": "new",
                "subject": "Hi",
                "body": "Hello",
            })

    def test_reply_requires_message_id(self):
        with pytest.raises(ValidationError, match="requires in_reply_to_message_id"):
            ComposeEmail.model_validate({
                "type": "email",
                "mode": "reply",
                "body": "Reply body",
            })

    def test_reply_with_message_id_valid(self):
        m = ComposeEmail.model_validate({
            "type": "email",
            "mode": "reply",
            "in_reply_to_message_id": "AAM123",
            "body": "Reply body",
        })
        assert m.mode == "reply"

    def test_reply_all_requires_message_id(self):
        with pytest.raises(ValidationError, match="requires in_reply_to_message_id"):
            ComposeEmail.model_validate({
                "type": "email",
                "mode": "reply_all",
                "body": "Reply body",
            })

    def test_forward_requires_message_id(self):
        with pytest.raises(ValidationError, match="requires in_reply_to_message_id"):
            ComposeEmail.model_validate({
                "type": "email",
                "mode": "forward",
                "body": "Forwarded content",
            })

    def test_extra_field_forbidden(self):
        with pytest.raises(ValidationError):
            ComposeEmail.model_validate({
                "type": "email",
                "mode": "new",
                "to": [{"email": "a@b.com"}],
                "subject": "Hi",
                "body": "Hello",
                "not_a_real_field": "boom",
            })

    def test_empty_body_rejected(self):
        with pytest.raises(ValidationError):
            ComposeEmail.model_validate({
                "type": "email",
                "mode": "new",
                "to": [{"email": "a@b.com"}],
                "subject": "Hi",
                "body": "",
            })


class TestComposeEvent:
    def test_valid_event(self):
        m = ComposeEvent.model_validate({
            "type": "event",
            "subject": "Sync",
            "start": "2026-07-08T10:00:00Z",
            "end": "2026-07-08T11:00:00Z",
            "timezone": "Europe/Berlin",
        })
        assert m.subject == "Sync"

    def test_end_before_start_rejected(self):
        with pytest.raises(ValidationError, match="end must be after start"):
            ComposeEvent.model_validate({
                "type": "event",
                "subject": "Sync",
                "start": "2026-07-08T11:00:00Z",
                "end": "2026-07-08T10:00:00Z",
                "timezone": "UTC",
            })

    def test_end_equals_start_rejected(self):
        with pytest.raises(ValidationError, match="end must be after start"):
            ComposeEvent.model_validate({
                "type": "event",
                "subject": "Sync",
                "start": "2026-07-08T10:00:00Z",
                "end": "2026-07-08T10:00:00Z",
                "timezone": "UTC",
            })

    def test_duration_over_12h_rejected(self):
        with pytest.raises(ValidationError, match="exceeds 12 hours"):
            ComposeEvent.model_validate({
                "type": "event",
                "subject": "Marathon",
                "start": "2026-07-08T00:00:00Z",
                "end": "2026-07-09T00:00:01Z",
                "timezone": "UTC",
            })

    def test_duration_exactly_12h_ok(self):
        m = ComposeEvent.model_validate({
            "type": "event",
            "subject": "Long",
            "start": "2026-07-08T00:00:00Z",
            "end": "2026-07-08T12:00:00Z",
            "timezone": "UTC",
        })
        assert m.subject == "Long"


class TestComposeTeamsMessage:
    def test_valid(self):
        m = ComposeTeamsMessage.model_validate({
            "type": "teams_message",
            "chat_id": "19:abc@thread.v2",
            "content": "hi",
        })
        assert m.content_type == "text"

    def test_content_type_literal(self):
        with pytest.raises(ValidationError):
            ComposeTeamsMessage.model_validate({
                "type": "teams_message",
                "chat_id": "19:abc@thread.v2",
                "content": "hi",
                "content_type": "markdown",
            })

    def test_html_content_type_ok(self):
        m = ComposeTeamsMessage.model_validate({
            "type": "teams_message",
            "chat_id": "19:abc@thread.v2",
            "content": "<b>hi</b>",
            "content_type": "html",
        })
        assert m.content_type == "html"

    def test_empty_chat_id_rejected(self):
        with pytest.raises(ValidationError):
            ComposeTeamsMessage.model_validate({
                "type": "teams_message",
                "chat_id": "",
                "content": "hi",
            })
