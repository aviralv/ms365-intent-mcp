# ms365-intent-mcp

An intent-oriented MCP server for Microsoft 365.

Most Microsoft 365 MCP servers replicate Graph API endpoints as tools — one tool per endpoint, orchestration pushed to the LLM. This server takes a different approach: tools map to user intents, not API structure. One tool call answers one question.

## Tools

| Tool | What it does |
|------|-------------|
| `resolve` | Takes any M365 URL (Teams message, email, file, meeting) and returns its content |
| `my_day` | Your day at a glance: calendar with bodies, mail summary, Teams activity |
| `meeting` | Everything about a meeting: metadata, body, attendees, transcript, chat |
| `compose` | Create email drafts, calendar events, or Teams messages |
| `whats_new` | Unified activity feed across calendar, mail, and Teams for a time window |
| `find` | Search across all M365 data (mail, files, SharePoint, Teams) |
| `people` | Person lookup with recent interaction context |
| `schedule` | Find available meeting times using the scheduling assistant |

## v0.8.0 dual-registered surface (interim)

As of v0.8.0, every tool is registered TWICE on the server:

- **Legacy names** (`my_day`, `meeting`, `compose`, `schedule`, `people`, `whats_new`, `find`, `resolve`) — the flat-kwargs shape from v0.7.0 and earlier. Still works identically. Emits a server-side deprecation warning on every call.
- **New v1 names** (`my_day_v1`, `meeting_v1`, `compose_v1`, `schedule_v1`, `people_v1`, `whats_new_v1`, `find_v1`, `resolve_v1`) — Pydantic `payload={...}` shape, typed responses with `type` discriminator, `rendered_markdown` field, `ErrorResponse` envelope, idempotency keys on writes.

**Migration timeline:**
1. **v0.8.0 (now)** — dual-registered. Existing consumers keep working. Deprecation warnings surface which callers still use legacy.
2. **v0.9.x (multiple releases)** — consumers of ms365-intent-mcp migrate from legacy to `_v1` names, one at a time.
3. **v1.0.0 (later)** — after every consumer has migrated and server logs show zero deprecation warnings for 7 consecutive days, legacy names are removed. The `_v1` variants are renamed to their canonical names (`my_day_v1` → `my_day`, etc.).

**For adopters starting fresh**: skip the legacy names entirely. Use `_v1` from the start; they will be renamed to canonical (unsuffixed) in v1.0.0.

## Design Philosophy

- **8 tools, not 76** — lean schema, loadable in any LLM session without bloat
- **Composed internally** — each tool makes 2-10 parallel Graph calls and fuses results
- **Decision-ready responses** — enough content to decide whether to engage, not just metadata
- **Permission-aware** — discovers available scopes at startup, degrades gracefully
- **No admin consent required** — device code flow with user-consentable scopes

## Installation

```bash
uv tool install ms365-intent-mcp
```

## Authentication

```bash
ms365-intent-mcp auth
```

Uses device code flow — visit the URL, enter the code, done. No app registration needed.

## Compliance

This tool accesses your Microsoft 365 data using your own credentials. You are responsible for ensuring usage complies with your organization's policies. The tool providers accept no liability for data handling — all data stays local and is processed on-device.

No data is sent to third parties. Authentication tokens are stored locally.

## License

MIT
