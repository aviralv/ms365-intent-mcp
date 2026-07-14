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
| `transcript` | Download a meeting recording's VTT (by URL, meeting name, or item_id+drive_id+site_root), or `list=true` to enumerate recordings newest-first |

## Design Philosophy

- **9 tools, not 76** — lean schema, loadable in any LLM session without bloat
- **Composed internally** — each tool makes 2-10 parallel Graph calls and fuses results
- **Decision-ready responses** — enough content to decide whether to engage, not just metadata
- **Permission-aware** — discovers available scopes at startup, degrades gracefully
- **No admin consent required** — device code flow with user-consentable scopes

## Installation

Not published to PyPI — install from GitHub:

```bash
uv tool install "ms365-intent-mcp[auth] @ git+https://github.com/aviralv/ms365-intent-mcp.git"
```

The `[auth]` extra pulls in `msal`, required by the `auth` subcommand below. To install a specific branch or tag, append `@<ref>` to the URL:

```bash
uv tool install "ms365-intent-mcp[auth] @ git+https://github.com/aviralv/ms365-intent-mcp.git@main"
```

To upgrade after a new release:

```bash
uv cache clean ms365-intent-mcp && uv tool install --force "ms365-intent-mcp[auth] @ git+https://github.com/aviralv/ms365-intent-mcp.git"
```

## Authentication

```bash
ms365-intent-mcp auth
```

Uses device code flow — visit the URL, enter the code, done. No app registration needed. Token is saved to `~/.config/ms365-intent-mcp/token.json` (0600).

## Compliance

This tool accesses your Microsoft 365 data using your own credentials. You are responsible for ensuring usage complies with your organization's policies. The tool providers accept no liability for data handling — all data stays local and is processed on-device.

No data is sent to third parties. Authentication tokens are stored locally.

## License

MIT
