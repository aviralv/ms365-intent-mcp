"""compose composer — dispatches to email_draft, reply_draft, event, teams_message."""

import html
import urllib.parse
from enum import Enum

from ..resolver import normalize_message_id

from ..formatters import (
    format_draft_created_markdown,
    format_event_created_markdown,
    format_event_forwarded_markdown,
)
from ..graph import GraphClient
from ..permissions import PermissionRegistry


class ComposeType(str, Enum):
    EMAIL_DRAFT = "email_draft"
    REPLY_DRAFT = "reply_draft"
    EMAIL_FORWARD = "email_forward"
    EVENT = "event"
    EVENT_FORWARD = "event_forward"
    TEAMS_MESSAGE = "teams_message"


SCOPE_REQUIREMENTS = {
    ComposeType.EMAIL_DRAFT: "Mail.ReadWrite",
    ComposeType.REPLY_DRAFT: "Mail.ReadWrite",
    ComposeType.EMAIL_FORWARD: "Mail.ReadWrite",
    ComposeType.EVENT: "Calendars.ReadWrite",
    ComposeType.EVENT_FORWARD: "Calendars.ReadWrite",
    ComposeType.TEAMS_MESSAGE: "ChatMessage.Send",
}


async def compose_action(
    client: GraphClient,
    permissions: PermissionRegistry,
    action_type: ComposeType,
    params: dict,
) -> tuple[dict, str]:
    required_scope = SCOPE_REQUIREMENTS[action_type]
    scope_msg = permissions.check(required_scope)
    if scope_msg:
        return {}, scope_msg

    if action_type == ComposeType.EMAIL_DRAFT:
        return await _create_email_draft(client, params)
    elif action_type == ComposeType.REPLY_DRAFT:
        return await _create_reply_draft(client, params)
    elif action_type == ComposeType.EMAIL_FORWARD:
        return await _forward_email_draft(client, params)
    elif action_type == ComposeType.EVENT:
        return await _create_event(client, params)
    elif action_type == ComposeType.EVENT_FORWARD:
        return await _forward_event(client, params)
    elif action_type == ComposeType.TEAMS_MESSAGE:
        return await _send_teams_message(client, params)
    else:
        return {}, f"❌ Unknown compose type: {action_type}"


async def _create_email_draft(client: GraphClient, params: dict) -> tuple[dict, str]:
    if not params.get("to"):
        return {}, "❌ Missing required field: 'to' (recipients list)"
    if not params.get("subject"):
        return {}, "❌ Missing required field: 'subject'"
    if not params.get("body"):
        return {}, "❌ Missing required field: 'body'"
    recipients = [
        {"emailAddress": {"address": r["email"], "name": r.get("name", r["email"])}}
        for r in params["to"]
    ]
    payload = {
        "subject": params["subject"],
        "body": {"contentType": "HTML", "content": params["body"]},
        "toRecipients": recipients,
    }
    if params.get("cc"):
        payload["ccRecipients"] = [
            {"emailAddress": {"address": r["email"], "name": r.get("name", r["email"])}}
            for r in params["cc"]
        ]
    if params.get("importance"):
        payload["importance"] = params["importance"]

    draft = await client.post("/me/messages", payload)
    data = {
        "draft_id": draft.get("id", ""),
        "subject": draft.get("subject", params.get("subject", "")),
        "to": [
            {"email": r.get("emailAddress", {}).get("address", ""), "name": r.get("emailAddress", {}).get("name", "")}
            for r in draft.get("toRecipients", [])
        ],
        "web_link": draft.get("webLink", ""),
    }
    return data, format_draft_created_markdown(draft)


async def _create_reply_draft(client: GraphClient, params: dict) -> tuple[dict, str]:
    if not params.get("message_id"):
        return {}, "❌ Missing required field: 'message_id'"
    if not params.get("body"):
        return {}, "❌ Missing required field: 'body'"
    message_id = params["message_id"]
    reply_all = params.get("reply_all", True)
    endpoint = "createReplyAll" if reply_all else "createReply"

    payload = {
        "message": {
            "body": {"contentType": "HTML", "content": params["body"]},
        }
    }
    if params.get("comment"):
        payload["comment"] = params["comment"]

    quoted_id = urllib.parse.quote(normalize_message_id(message_id), safe="")
    draft = await client.post(f"/me/messages/{quoted_id}/{endpoint}", payload)
    data = {
        "draft_id": draft.get("id", ""),
        "subject": draft.get("subject", ""),
        "to": [
            {"email": r.get("emailAddress", {}).get("address", ""), "name": r.get("emailAddress", {}).get("name", "")}
            for r in draft.get("toRecipients", [])
        ],
        "web_link": draft.get("webLink", ""),
    }
    return data, format_draft_created_markdown(draft)


async def _forward_email_draft(client: GraphClient, params: dict) -> tuple[dict, str]:
    if not params.get("message_id"):
        return {}, "❌ Missing required field: 'message_id'"
    if not params.get("to"):
        return {}, "❌ Missing required field: 'to' (recipients list)"
    message = {
        "toRecipients": [
            {"emailAddress": {"address": r["email"], "name": r.get("name", r["email"])}}
            for r in params["to"]
        ],
    }
    if params.get("cc"):
        message["ccRecipients"] = [
            {"emailAddress": {"address": r["email"], "name": r.get("name", r["email"])}}
            for r in params["cc"]
        ]
    if params.get("body"):
        message["body"] = {"contentType": "HTML", "content": html.escape(params["body"])}

    draft = await client.post(
        f"/me/messages/{urllib.parse.quote(normalize_message_id(params['message_id']), safe='')}/createForward",
        {"message": message},
    )
    data = {
        "draft_id": draft.get("id", ""),
        "subject": draft.get("subject", ""),
        "to": [
            {
                "email": r.get("emailAddress", {}).get("address", ""),
                "name": r.get("emailAddress", {}).get("name", ""),
            }
            for r in draft.get("toRecipients", [])
        ],
        "web_link": draft.get("webLink", ""),
    }
    return data, format_draft_created_markdown(draft)


async def _create_event(client: GraphClient, params: dict) -> tuple[dict, str]:
    if not params.get("subject"):
        return {}, "❌ Missing required field: 'subject'"
    if not params.get("start"):
        return {}, "❌ Missing required field: 'start'"
    if not params.get("end"):
        return {}, "❌ Missing required field: 'end'"
    tz = params.get("timezone", "UTC")
    payload = {
        "subject": params["subject"],
        "start": {"dateTime": params["start"], "timeZone": tz},
        "end": {"dateTime": params["end"], "timeZone": tz},
    }
    if params.get("body"):
        payload["body"] = {"contentType": "HTML", "content": params["body"]}
    if params.get("location"):
        payload["location"] = {"displayName": params["location"]}
    if params.get("attendees"):
        payload["attendees"] = [
            {
                "emailAddress": {"address": a["email"], "name": a.get("name", a["email"])},
                "type": a.get("type", "required"),
            }
            for a in params["attendees"]
        ]
    if params.get("is_online_meeting"):
        payload["isOnlineMeeting"] = True

    event = await client.post("/me/events", payload)
    join_url = (event.get("onlineMeeting") or {}).get("joinUrl", "") if event.get("isOnlineMeeting") else ""
    data = {
        "event_id": event.get("id", ""),
        "subject": event.get("subject", params.get("subject", "")),
        "start": (event.get("start") or {}).get("dateTime", params.get("start", "")),
        "end": (event.get("end") or {}).get("dateTime", params.get("end", "")),
        "join_url": join_url or None,
    }
    return data, format_event_created_markdown(event)


async def _forward_event(client: GraphClient, params: dict) -> tuple[dict, str]:
    if not params.get("event_id"):
        return {}, "❌ Missing required field: 'event_id'"
    if not params.get("to"):
        return {}, "❌ Missing required field: 'to' (recipients list)"
    body: dict = {
        "ToRecipients": [
            {"EmailAddress": {"Address": r["email"], "Name": r.get("name", r["email"])}}
            for r in params["to"]
        ],
    }
    if params.get("comment"):
        body["Comment"] = params["comment"]

    await client.post(f"/me/events/{urllib.parse.quote(normalize_message_id(params['event_id']), safe='')}/forward", body)
    to_out = [{"email": r["email"], "name": r.get("name", r["email"])} for r in params["to"]]
    to_names = [r["name"] for r in to_out]
    data = {"to": to_out}
    return data, format_event_forwarded_markdown(to_names, params.get("comment"))


async def _send_teams_message(client: GraphClient, params: dict) -> tuple[dict, str]:
    if not params.get("chat_id"):
        return {}, "❌ Missing required field: 'chat_id'"
    if not params.get("content"):
        return {}, "❌ Missing required field: 'content'"
    chat_id = params["chat_id"]
    content = params["content"]
    content_type = params.get("content_type", "text")

    payload = {
        "body": {"contentType": content_type, "content": content},
    }
    msg = await client.post(f"/chats/{chat_id}/messages", payload)
    data = {
        "message_id": (msg or {}).get("id", ""),
        "chat_id": chat_id,
    }
    return data, "✅ Message sent to Teams chat."
