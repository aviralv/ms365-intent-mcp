"""resolve composer — parse M365 URLs and fetch their content via Graph."""

from ..formatters import format_resolved_content_markdown, format_section_error
from ..graph import GraphAPIError, GraphClient
from ..permissions import PermissionRegistry
from ..resolver import ResolvedUrl, UrlParseError, resolve_url
from ._utils import _error_reason


async def compose_resolve(
    client: GraphClient,
    permissions: PermissionRegistry,
    url: str,
) -> str:
    try:
        resolved = resolve_url(url)
    except UrlParseError as exc:
        return f"❌ Unrecognised URL — {exc}"

    scope_msg = permissions.check(resolved.required_scope)
    if scope_msg:
        return scope_msg

    try:
        data = await _fetch_resolved(client, resolved)
    except GraphAPIError as exc:
        return format_section_error("Resolve", _error_reason(exc))

    return format_resolved_content_markdown(resolved.url_type, data)


async def _fetch_resolved(client: GraphClient, resolved: ResolvedUrl) -> dict:
    url_type = resolved.url_type
    endpoint = resolved.graph_endpoint

    if url_type == "email":
        return await client.get(endpoint, params={
            "$select": "subject,from,receivedDateTime,bodyPreview,body",
        })
    elif url_type == "channel_message":
        result = await client.get(endpoint)
        messages = (result or {}).get("value", [])
        return messages[0] if messages else {}
    elif url_type == "chat_message":
        result = await client.get(endpoint, params={"$top": "1"})
        messages = (result or {}).get("value", [])
        return messages[0] if messages else {}
    elif url_type == "meeting":
        return {"subject": "Meeting", "body": {"content": "Meeting details via Teams link."}}
    elif url_type == "sharepoint_page":
        return await client.get(endpoint)
    elif url_type in ("onedrive_file", "onedrive_share_link"):
        return await client.get("/me/drive/root", params={"$select": "name,size,webUrl"})
    else:
        return {}
