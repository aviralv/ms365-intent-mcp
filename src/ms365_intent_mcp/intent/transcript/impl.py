"""transcript implementation — wraps composers.transcript.compose_transcript."""

from __future__ import annotations

from fastmcp import Context

from ...composers.transcript import compose_transcript
from .._helpers import _get_deps, wrap_errors
from .schemas import TranscriptPayload, TranscriptResultResponse

TOOL_NAME = "transcript"


@wrap_errors(TOOL_NAME)
async def _transcript_impl(ctx: Context, payload: TranscriptPayload) -> TranscriptResultResponse:
    """Call the underlying composer and return a typed TranscriptResultResponse."""
    _config, client, permissions = _get_deps(ctx)
    vroom = ctx.request_context.lifespan_context["vroom"]

    data, markdown = await compose_transcript(
        client,
        vroom,
        permissions,
        url=payload.url,
        name=payload.name,
        item_id=payload.item_id,
        drive_id=payload.drive_id,
        site_root=payload.site_root,
        output_dir=payload.output_dir,
        list_recordings=payload.list,
    )

    return TranscriptResultResponse(
        status=data.get("status", "error"),
        file_path=data.get("file_path") or None,
        meeting_name=data.get("meeting_name") or None,
        meeting_date=data.get("meeting_date") or None,
        line_count=data.get("line_count", 0),
        byte_count=data.get("byte_count", 0),
        has_speaker_tags=data.get("has_speaker_tags", False),
        alternatives_count=data.get("alternatives_count", 0),
        recordings=data.get("recordings"),
        message=data.get("message") or None,
        rendered_markdown=markdown,
    )
