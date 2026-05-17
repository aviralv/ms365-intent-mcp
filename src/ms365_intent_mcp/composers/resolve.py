"""resolve composer — parse M365 URLs and fetch their content via Graph."""

from datetime import datetime, timedelta, timezone

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

    if "_error" in data:
        return format_section_error("Resolve", data["_error"])

    return format_resolved_content_markdown(resolved.url_type, data)


async def _fetch_resolved(client: GraphClient, resolved: ResolvedUrl) -> dict:
    url_type = resolved.url_type
    endpoint = resolved.graph_endpoint

    if url_type == "email":
        return await client.get(endpoint, params={
            "$select": "subject,from,receivedDateTime,bodyPreview,body",
        })

    elif url_type == "channel_message":
        return await client.get(endpoint, params={
            "$select": "body,from,createdDateTime,subject",
        })

    elif url_type == "chat_message":
        return await client.get(endpoint, params={
            "$select": "body,from,createdDateTime",
        })

    elif url_type == "meeting":
        thread_id = resolved.extra.get("thread_id", "")
        now = datetime.now(timezone.utc)
        start = (now - timedelta(days=14)).strftime("%Y-%m-%dT%H:%M:%SZ")
        end = (now + timedelta(days=14)).strftime("%Y-%m-%dT%H:%M:%SZ")
        result = await client.get("/me/calendarView", params={
            "startDateTime": start,
            "endDateTime": end,
            "$top": "50",
            "$select": "subject,start,end,organizer,attendees,body,location,isOnlineMeeting,onlineMeeting",
        })
        events = (result or {}).get("value", [])
        for event in events:
            join_url = (event.get("onlineMeeting") or {}).get("joinUrl", "")
            if thread_id and thread_id in join_url:
                return event
        return {"_error": "No matching meeting found for this Teams link."}

    elif url_type in ("onedrive_file", "onedrive_share_link"):
        return await client.get(endpoint, params={
            "$select": "name,size,webUrl,lastModifiedDateTime,createdDateTime,file",
        })

    elif url_type == "sharepoint_page":
        site_data = await client.get(endpoint)
        site_id = (site_data or {}).get("id", "")
        page_filename = resolved.extra.get("page_filename", "")
        if site_id and page_filename:
            try:
                page_data = await _fetch_sharepoint_page(client, site_id, page_filename)
                if page_data:
                    page_data["_page_found"] = True
                    page_data["_site_name"] = (site_data or {}).get("displayName", "")
                    return page_data
            except GraphAPIError:
                pass
        return site_data

    else:
        return {}


async def _fetch_sharepoint_page(client: GraphClient, site_id: str, filename: str) -> dict | None:
    """Look up a SharePoint page by filename via the Site Pages list."""
    lists_result = await client.get(
        f"/sites/{site_id}/lists",
        params={"$filter": "displayName eq 'Site Pages'", "$select": "id"},
    )
    lists = (lists_result or {}).get("value", [])
    if not lists:
        return None
    list_id = lists[0]["id"]
    items_result = await client.get(
        f"/sites/{site_id}/lists/{list_id}/items",
        params={
            "$filter": f"fields/FileLeafRef eq '{filename}'",
            "$select": "id,webUrl",
            "$expand": "fields($select=FileLeafRef,Title,Modified)",
            "$top": "1",
        },
        headers={"Prefer": "HonorNonIndexedQueriesWarningMayFailRandomly"},
    )
    items = (items_result or {}).get("value", [])
    if not items:
        return None
    item = items[0]
    fields = item.get("fields", {})
    return {
        "name": fields.get("FileLeafRef", filename),
        "title": fields.get("Title", ""),
        "webUrl": item.get("webUrl", ""),
        "lastModifiedDateTime": fields.get("Modified", ""),
    }
