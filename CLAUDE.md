# ms365-intent-mcp

Intent-oriented MCP server for Microsoft 365. Tools match how you think about your work day, not how Microsoft organizes its Graph API.

## Project Intent

Replace 76 API-mirroring tools with ~8 composed intent tools. Eliminate the need for separate email/teams-chat/onedrive routing agents. Schema lean enough to load in the base Claude session.

## Tech Stack

- Python 3.11+
- FastMCP (MCP server framework)
- httpx (async HTTP client)
- pydantic-settings (configuration)
- MSAL (auth script only, not runtime)

## Architecture

```
Tool Layer (8 intent tools) → thin validation + formatting
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
- MVP scopes (SAP-verified): Calendars.ReadWrite, Mail.Read/ReadWrite, Files.Read, Sites.Read.All, Chat.ReadWrite, ChatMessage.Send, Channel.ReadBasic.All, Team.ReadBasic.All, Tasks.Read, Contacts.Read, User.Read

## Tool Surface

| Tool | Intent |
|------|--------|
| `resolve(url)` | "What is this link?" — any M365 URL → content |
| `my_day(date?)` | "What does my day look like?" — calendar + mail + teams |
| `meeting(identifier)` | "Tell me about this meeting" — full context |
| `compose(type, params)` | "Create something" — drafts, events, messages |
| `whats_new(since, scope?)` | "What happened?" — unified activity |
| `find(query, type?)` | "Find me X" — cross-domain search |
| `people(query)` | "Who is this?" — person + interaction context |
| `schedule(attendees, duration)` | "Find a time" — scheduling assistant |

## Phasing

- **Phase 1 (MVP)**: resolve, my_day, meeting, compose
- **Phase 2**: whats_new, find, people, schedule
- **Phase 3**: Knowledge layer (MCP resources), generalization (multi-tenant)
