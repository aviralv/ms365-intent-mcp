# ms365-intent-mcp

## Tech Stack

- Python 3.11+
- FastMCP (MCP server framework)
- httpx (async HTTP client)
- pydantic-settings (configuration)
- MSAL (auth script only, not runtime)

## Architecture

```
Tool Layer (intent tools) → thin validation + formatting
Composer Layer → orchestrates parallel Graph calls, fuses results
Graph Client → httpx with retry/backoff, rate limit awareness
Auth Layer → device code flow, silent token refresh
Permission Registry → scope introspection, graceful degradation
```

Each tool has a composer that makes 2-10 parallel Graph calls and returns a fused, decision-ready response.

## Development Workflow

```bash
# Install locally
uv tool install --force .

# Reinstall after changes (bump version in pyproject.toml first)
uv cache clean ms365-intent-mcp && uv tool install --force .

# Run auth
uv run python scripts/authenticate.py

# Test directly
uv run python -m ms365_intent_mcp.server
```

## Design Constraints

- 8 tools max (schema must be lean)
- Parallel Graph calls by default (asyncio.gather)
- Partial success: if one subsystem fails, others still return
- Decision-ready responses: include enough content to decide whether to engage
- No raw IDs in markdown responses — names, subjects, links only
- Markdown output only (no JSON format unless a consumer emerges)
- MVP scopes (SAP-verified): Calendars.ReadWrite, Mail.Read/ReadWrite, Files.Read, Sites.Read.All, Chat.ReadWrite, ChatMessage.Send, Channel.ReadBasic.All, Team.ReadBasic.All, Tasks.Read, Contacts.Read, User.Read

## Coding Discipline

- State assumptions before implementing. If multiple interpretations exist, surface them — don't pick silently.
- Minimum code that solves the problem. No speculative features, no premature abstractions, no "flexibility" that wasn't requested.
- Surgical changes: every changed line traces to the task. Don't "improve" adjacent code, comments, or formatting.
- Your mess → clean up (imports/variables your changes made unused). Pre-existing mess → mention it, don't touch it.
- Multi-step tasks: state a brief plan with verification at each step before starting.
- The overcomplicated version isn't wrong — it's premature. Solve today's problem simply; refactor when complexity is actually needed.

## Session Notes

Session notes live in the orchestration directory, not here:
`the-product-kitchen/session-notes/ms365-intent-mcp-YYYY-MM-DD-theme.md`
