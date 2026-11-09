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

### Occurrence-safe meeting lookup

```python
# Synthetic examples; identifiers below are not real Graph IDs.
meeting(payload={"identifier": "Synthetic Weekly Sync", "date": "2026-11-09"})
my_day(payload={"date": "2026-11-09", "include_bodies": True})
# Pass an events[i].id from the structured my_day response:
meeting(payload={"identifier": "synthetic-occurrence-id"})
```

- **Explicit `date`** must be `YYYY-MM-DD` and selects an occurrence whose **start date** is that day in the configured calendar timezone (`default_timezone`, UTC by default). The lookup uses Graph `calendarView` to expand recurring events, with timezone-offset day bounds and a next-midnight boundary that handles DST. Events overlapping from a previous day are not selected. If several subject-substring matches start that day, the earliest is selected; use the event ID to disambiguate further.
- **Date misses** return `ErrorResponse(code="not_found")`, never a different-date occurrence. A valid event ID with a date is checked against that date without subject fallback; a series-master ID with a date is rejected (`validation_error`), so use an occurrence ID from `my_day`. `date` cannot be combined with `"next"`. Unsupported configured timezones are rejected rather than guessed.
- **Compatibility:** omitting `date` preserves the existing ID / `"next"` / subject lookup. In particular, subject-only lookup still selects the first ascending match in the existing ±7-day window; it does **not** default to today. A date suffix inside `identifier` is still a literal part of the subject search, not a date parameter.
- **Structured IDs:** `my_day.events[].id` contains the Graph event/occurrence ID with or without bodies. IDs are not added to rendered human markdown. Existing event fields are unchanged; `id` defaults to `""` when parsing older data without an ID. Explicit-date subject lookup follows up to five pages (50 events/page); an unresolved truncated lookup returns `graph_api_error`, not a definitive date miss. The existing `my_day` 30-event limit is unchanged.
- **Recording guard:** `recording.occurrence_date` and `date_matches_occurrence` still come from the real selected event. If only another occurrence's recording exists, it remains visibly flagged stale (`False`); no requested-date override is applied. The existing event-timezone vs. Teams filename tenant-local-date boundary caveat remains: a cross-midnight timezone difference can cause a conservative false-stale warning.

### Notes on `transcript`

- **Recordings made by another attendee** live on their drive and are invisible to `name`/`list` discovery. Copy the recording's link from the Teams Recap UI and pass it via `url` — it carries the driveId+driveItemId that resolve the item directly, and file access survives even after you lose access to the meeting chat.
- **Ad-hoc 1:1 calls** are titled from the recorder's side, so searching the counterpart's name may miss them — use `list=true` to find them by date.
- **`name` discovery** searches own-drive, Graph Search, and Teams chats; on multiple matches the most recent is downloaded and the others are surfaced.

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
