"""people composer — person lookup with parallel email + Teams context."""

import asyncio

from ..formatters import format_people_markdown
from ..graph import GraphClient, GraphAPIError
from ..permissions import PermissionRegistry
from ._utils import _escape_odata, _list_user_chats, _prefilter_chats_by_query


async def compose_people(
    client: GraphClient,
    permissions: PermissionRegistry,
    query: str,
) -> tuple[dict, str]:
    people = await _lookup_person(client, permissions, query)
    if not people:
        markdown = f"### People\nNo results for '{query}'."
        data: dict = {
            "name": query,
            "email": "",
            "job_title": None,
            "recent_mail": [],
            "recent_chat": None,
        }
        return data, markdown

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
        if emails_result and not isinstance(emails_result, BaseException):
            recent_emails = (emails_result or {}).get("value", [])

        chats_result = results.get("chats")
        if chats_result and not isinstance(chats_result, BaseException):
            chats = (chats_result or {}).get("value", [])
            recent_chat = _find_chat_with_person(chats, display_name, email_addr)

    markdown = format_people_markdown(query, people, recent_emails, recent_chat)
    data = {
        "name": display_name or query,
        "email": email_addr,
        "job_title": person.get("jobTitle") or None,
        "recent_mail": [
            {
                "subject": m.get("subject", ""),
                "sender": (m.get("from") or {}).get("emailAddress", {}).get("name", ""),
                "received": m.get("receivedDateTime"),
            }
            for m in recent_emails
        ],
        "recent_chat": {
            "body": ((recent_chat.get("lastMessagePreview") or {}).get("body") or {}).get("content", ""),
            "last_message_at": (recent_chat.get("lastMessagePreview") or {}).get("createdDateTime"),
            "chat_id": recent_chat.get("id", ""),
            "chat_url": recent_chat.get("webUrl", ""),
        } if recent_chat else None,
    }
    return data, markdown


def _find_chat_with_person(
    chats: list[dict], display_name: str, email: str
) -> dict | None:
    """Find the most recent chat that includes the target person.

    Prefers email match (authoritative). Falls back to all-words-match on
    displayName so a short query ('Alice') doesn't match longer names
    that happen to contain those characters ('Alice Patel-Smith').
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


async def _get_me_id(client: GraphClient) -> str:
    """Return the signed-in user's AAD object id, or '' on failure.

    Used for self-exclusion in the chat-membership fallback: a chat member's
    `userId` is the same AAD object id, so equality is reliable (unlike
    UPN-vs-email comparison, which diverges in real tenants).
    """
    try:
        me = await client.get("/me", params={"$select": "id"})
    except GraphAPIError:
        return ""
    return (me or {}).get("id", "")


def _lookup_person_via_chats(
    chats: list[dict], query: str, me_id: str
) -> list[dict]:
    """Synthesize matched people from chat members (pure — no API call).

    Third fallback tier for compose_people: used when /me/people and
    /me/contacts both miss. Matches a member when ALL query words are a subset
    of the member's displayName words (case-insensitive), so 'Avi' does not
    match 'Aviral'. The signed-in user is excluded by `me_id` (their member
    `userId`). Results are deduped by userId → email → displayName and ordered
    by chat recency (chats arrive newest-first).
    """
    narrowed, _ = _prefilter_chats_by_query(chats, query)
    target_words = {w for w in query.lower().split() if w}
    if not target_words:
        return []

    me_id_lower = (me_id or "").lower()
    seen: set[str] = set()
    people: list[dict] = []
    for chat in narrowed:
        for member in chat.get("members", []):
            user_id = (member.get("userId") or "")
            if me_id_lower and user_id.lower() == me_id_lower:
                continue
            display_name = member.get("displayName") or ""
            member_words = set(display_name.lower().split())
            if not target_words.issubset(member_words):
                continue
            email = member.get("email") or ""
            key = user_id.lower() or email.lower() or display_name.lower()
            if not key or key in seen:
                continue
            seen.add(key)
            people.append({
                "displayName": display_name,
                "emailAddresses": [{"address": email}] if email else [],
                "jobTitle": None,
                "_source_chat": chat,
            })
    return people


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
