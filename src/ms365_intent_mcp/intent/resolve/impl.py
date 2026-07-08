"""resolve_v1 implementation — dispatches URL to compose_resolve and wraps
the structured return into a typed ResolvedContent response."""

from __future__ import annotations

from fastmcp import Context

from ...composers.resolve import compose_resolve
from .._helpers import IntentError, _get_deps, wrap_errors
from .schemas import (
    ChannelMessageContent,
    ChatMessageContent,
    ChatThreadContent,
    EmailContent,
    MeetingContent,
    OneDriveFileContent,
    ResolvedContent,
    ResolvePayload,
    SharePointPageContent,
)

TOOL_NAME = "resolve_v1"

_KIND_TO_MODEL = {
    "email": EmailContent,
    "chat_thread": ChatThreadContent,
    "chat_message": ChatMessageContent,
    "channel_message": ChannelMessageContent,
    "meeting": MeetingContent,
    "sharepoint_page": SharePointPageContent,
    "onedrive_file": OneDriveFileContent,
    "onedrive_share_link": OneDriveFileContent,
}


@wrap_errors(TOOL_NAME)
async def _resolve_v1_impl(ctx: Context, payload: ResolvePayload) -> ResolvedContent:
    """Resolve an M365 URL and return a typed response."""
    _, client, permissions = _get_deps(ctx)

    from ...resolver import UrlParseError, resolve_url

    try:
        resolved = resolve_url(str(payload.url))
    except UrlParseError as exc:
        raise IntentError("invalid_id", str(exc))

    url_type = resolved.url_type
    data_dict, markdown = await compose_resolve(client, permissions, str(payload.url))

    kind = data_dict.get("kind", url_type)
    structured_data = data_dict.get("data", {})

    model_cls = _KIND_TO_MODEL.get(kind)
    if model_cls is None:
        raise IntentError("invalid_id", f"Unknown url_type: {kind}")

    try:
        content_obj = model_cls.model_validate(structured_data)
    except Exception:
        # Fallback: construct with just the kind field
        content_obj = model_cls(kind=content_obj.model_fields["kind"].default)  # type: ignore[attr-defined]

    # Normalise onedrive_share_link → onedrive_file for the kind discriminator
    canonical_kind = "onedrive_file" if kind == "onedrive_share_link" else kind

    return ResolvedContent(
        url=payload.url,
        kind=canonical_kind,
        data=content_obj,
        rendered_markdown=markdown,
    )
