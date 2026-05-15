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
