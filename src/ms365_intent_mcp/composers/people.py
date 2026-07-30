"""people composer — person lookup with chat-membership fallback."""

from __future__ import annotations

from ..formatters import format_people_markdown
from ..graph import GraphClient, GraphAPIError
from ..permissions import PermissionRegistry
from ._utils import _escape_odata, _list_user_chats, _prefilter_chats_by_query


def _extract_email(record: dict) -> str:
    """Return a usable email from any Graph person/contact/user shape, else ''.

    - contact / directory: emailAddresses[].address
    - person (relevance): scoredEmailAddresses[].address
    - /users directory: mail, else userPrincipalName if it is a real address
      (guest UPNs carry '#EXT#' and are not routable — rejected).
    """
    ea = record.get("emailAddresses") or []
    if ea and ea[0].get("address"):
        return ea[0]["address"]
    sea = record.get("scoredEmailAddresses") or []
    if sea and sea[0].get("address"):
        return sea[0]["address"]
    mail = record.get("mail")
    if mail:
        return mail
    upn = record.get("userPrincipalName") or ""
    if upn and "#EXT#" not in upn:
        return upn
    return ""


async def compose_people(
    client: GraphClient,
    permissions: PermissionRegistry,
    query: str,
) -> tuple[dict, str]:
    people = await _lookup_person(client, permissions, query)

    # Fetch the chat list once, up front: it powers both the person fallback
    # (when People/contacts miss) and recent_chat enrichment below.
    chats: list[dict] = []
    if permissions.has("Chat.ReadWrite"):
        try:
            chats = await _list_user_chats(client)
        except GraphAPIError:
            chats = []

    if not people and chats:
        me_id = await _get_me_id(client)
        people = _lookup_person_via_chats(chats, query, me_id)

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
    # Ambiguity fail-safe: >1 candidate from the resolving tier → withhold
    # the structured email (a forward would consume it). Single confident
    # hit → use its email. Email extraction handles all Graph shapes.
    if len(people) > 1:
        email_addr = ""
    else:
        email_addr = _extract_email(person)

    recent_emails: list[dict] = []
    if email_addr and permissions.has("Mail.Read"):
        try:
            emails_result = await client.get("/me/messages", params={
                "$filter": f"from/emailAddress/address eq '{_escape_odata(email_addr)}'",
                "$select": "subject,from,receivedDateTime",
                "$orderby": "receivedDateTime desc",
                "$top": "5",
            })
            recent_emails = (emails_result or {}).get("value", [])
        except GraphAPIError:
            recent_emails = []

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
    except Exception:
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
    escaped = _escape_odata(query)

    async def _people_tier() -> list[dict]:
        if not permissions.has("People.Read"):
            return []
        try:
            r = await client.get("/me/people", params={
                "$search": f'"{escaped}"', "$top": "5",
                "$select": "displayName,jobTitle,scoredEmailAddresses",
            })
            return (r or {}).get("value", [])
        except GraphAPIError:
            return []

    async def _contacts_tier() -> list[dict]:
        try:
            r = await client.get("/me/contacts", params={
                "$search": f'"{escaped}"', "$top": "5",
                "$select": "displayName,emailAddresses,jobTitle",
            })
            return (r or {}).get("value", [])
        except GraphAPIError:
            return []

    async def _users_tier() -> list[dict]:
        if not permissions.has("User.ReadBasic.All"):
            return []
        try:
            r = await client.get("/users", params={
                "$search": f'"displayName:{escaped}"', "$top": "5",
                "$select": "displayName,mail,userPrincipalName,jobTitle",
            }, headers={"ConsistencyLevel": "eventual"})
            return (r or {}).get("value", [])
        except GraphAPIError:
            return []

    # First tier that yields at least one email-bearing record wins; an
    # email-less hit no longer short-circuits the cascade.
    for tier in (_people_tier, _contacts_tier, _users_tier):
        results = await tier()
        if any(_extract_email(p) for p in results):
            return results
    return []
