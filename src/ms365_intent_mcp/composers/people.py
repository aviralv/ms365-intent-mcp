"""people composer — person lookup with parallel email + Teams context."""

import asyncio

from ..formatters import format_people_markdown
from ..graph import GraphClient, GraphAPIError
from ..permissions import PermissionRegistry
from ._utils import _escape_odata


async def compose_people(
    client: GraphClient,
    permissions: PermissionRegistry,
    query: str,
) -> str:
    people = await _lookup_person(client, permissions, query)
    if not people:
        return f"### People\nNo results for '{query}'."

    person = people[0]
    display_name = person.get("displayName", "")
    email_addr = ""
    email_addresses = person.get("emailAddresses", [])
    if email_addresses:
        email_addr = email_addresses[0].get("address", "")

    tasks = {}
    if email_addr and permissions.has("Mail.Read"):
        tasks["emails"] = client.get("/me/messages", params={
            "$filter": f"from/emailAddress/address eq '{_escape_odata(email_addr)}'",
            "$select": "subject,from,receivedDateTime",
            "$orderby": "receivedDateTime desc",
            "$top": "5",
        })

    if permissions.has("Chat.ReadWrite"):
        tasks["chats"] = client.get("/me/chats", params={
            "$expand": "members,lastMessagePreview",
            "$top": "20",
        })

    recent_emails: list[dict] = []
    recent_chat: dict | None = None

    if tasks:
        keys = list(tasks.keys())
        results_list = await asyncio.gather(*tasks.values(), return_exceptions=True)
        results = dict(zip(keys, results_list))

        emails_result = results.get("emails")
        if emails_result and not isinstance(emails_result, Exception):
            recent_emails = (emails_result or {}).get("value", [])

        chats_result = results.get("chats")
        if chats_result and not isinstance(chats_result, Exception):
            chats = (chats_result or {}).get("value", [])
            recent_chat = _find_chat_with_person(chats, display_name, email_addr)

    return format_people_markdown(query, people, recent_emails, recent_chat)


def _find_chat_with_person(
    chats: list[dict], display_name: str, email: str
) -> dict | None:
    """Find the most recent chat that includes the target person.

    Prefers email match (authoritative). Falls back to all-words-match on
    displayName so 'Avi' doesn't accidentally match 'Aviral Patel'.
    """
    email_lower = email.lower()
    if email_lower:
        for chat in chats:
            for member in chat.get("members", []):
                if (member.get("email") or "").lower() == email_lower:
                    return chat

    target_words = {w for w in display_name.lower().split() if w}
    if not target_words:
        return None
    for chat in chats:
        for member in chat.get("members", []):
            member_words = set((member.get("displayName") or "").lower().split())
            if target_words.issubset(member_words):
                return chat
    return None


async def _lookup_person(
    client: GraphClient,
    permissions: PermissionRegistry,
    query: str,
) -> list[dict]:
    if permissions.has("People.Read"):
        try:
            result = await client.get("/me/people", params={
                "$search": query,
                "$top": "5",
                "$select": "displayName,jobTitle,emailAddresses,phones",
            })
            return (result or {}).get("value", [])
        except GraphAPIError:
            pass

    try:
        result = await client.get("/me/contacts", params={
            "$search": f'"{query}"',
            "$top": "5",
            "$select": "displayName,emailAddresses,jobTitle",
        })
        return (result or {}).get("value", [])
    except GraphAPIError:
        return []
