# ms365-intent-mcp — Project Context

Intent-oriented MCP server for Microsoft 365. Tools match how you think about your work day, not how Microsoft organizes its Graph API.

## Problem

Replace 76 API-mirroring tools with ~8 composed intent tools. Eliminate the need for separate email/teams-chat/onedrive routing agents. Schema lean enough to load in the base Claude session.

## Tool Surface

| Tool | Intent |
|------|--------|
| `my_day(date?)` | "What does my day look like?" — calendar + mail + teams |
| `meeting(identifier)` | "Tell me about this meeting" — full context |
| `compose(type, params)` | "Create something" — drafts, events, messages |
| `resolve(url)` | "What is this link?" — any M365 URL → content |
| `whats_new(since, scope?)` | "What happened?" — unified activity |
| `find(query, type?)` | "Find me X" — cross-domain search |
| `people(query)` | "Who is this?" — person + interaction context |
| `schedule(attendees, duration)` | "Find a time" — scheduling assistant |

## Phasing

- **Phase 1 (MVP)**: my_day, meeting, compose + infra (auth, graph client, permission registry)
- **Phase 2**: resolve, whats_new, find, people, schedule + pagination support
- **Phase 3**: Knowledge layer (MCP resources), generalization (multi-tenant)

## References

- **Design spec** (detailed decisions, auth research, response patterns): `docs/specs/2026-05-15-ms365-intent-mcp-design.md`
- **Existing server** (reference for graph client, auth, resilience patterns): `../microsoft-365-mcp/`
- **Community servers** (auth patterns to study): `softeria/ms-365-mcp-server`, `pnp/cli-microsoft365-mcp-server` on GitHub
