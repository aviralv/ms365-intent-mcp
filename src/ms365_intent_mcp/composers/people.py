"""people composer — person lookup with parallel email + Teams context."""

import asyncio

from ..formatters import format_people_markdown
from ..graph import GraphClient, GraphAPIError
from ..permissions import PermissionRegistry


async def compose_people(
    client: GraphClient,
    permissions: PermissionRegistry,
    query: str,
) -> str:
    people = await _lookup_person(client, permissions, query)
    if not people:
        return f"### People\nNo results for '{query}'."

    person = people[0]
    email_addr = ""
    email_addresses = person.get("emailAddresses", [])
    if email_addresses:
        email_addr = email_addresses[0].get("address", "")

    tasks = {}
    if email_addr and permissions.has("Mail.Read"):
        tasks["emails"] = client.get("/me/messages", params={
            "$filter": f"from/emailAddress/address eq '{email_addr}'",
            "$select": "subject,from,receivedDateTime",
            "$orderby": "receivedDateTime desc",
            "$top": "5",
        })

    if permissions.has("Chat.ReadWrite"):
        tasks["chats"] = client.get("/me/chats", params={
            "$expand": "lastMessagePreview",
            "$top": "5",
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
            recent_chat = chats[0] if chats else None

    return format_people_markdown(query, people, recent_emails, recent_chat)


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

    # Fallback: contacts search
    try:
        result = await client.get("/me/contacts", params={
            "$search": f'"{query}"',
            "$top": "5",
            "$select": "displayName,emailAddresses,jobTitle",
        })
        return (result or {}).get("value", [])
    except GraphAPIError:
        return []
