"""Intent-oriented tool surface (v1) — payload={...} shape with typed responses.

``register_all(mcp)`` registers all 9 tools in one call.
"""

from __future__ import annotations

from fastmcp import FastMCP


def register_all(mcp: FastMCP) -> None:
    """Register every v1 tool on the FastMCP instance."""
    # Imports are inside the function to avoid module-import-order surprises
    # when this package is imported before FastMCP is fully initialized.
    from .compose import register as _r_compose
    from .find import register as _r_find
    from .meeting import register as _r_meeting
    from .my_day import register as _r_my_day
    from .people import register as _r_people
    from .resolve import register as _r_resolve
    from .schedule import register as _r_schedule
    from .transcript import register as _r_transcript
    from .whats_new import register as _r_whats_new

    _r_compose(mcp)
    _r_find(mcp)
    _r_meeting(mcp)
    _r_my_day(mcp)
    _r_people(mcp)
    _r_resolve(mcp)
    _r_schedule(mcp)
    _r_transcript(mcp)
    _r_whats_new(mcp)


__all__ = ["register_all"]
