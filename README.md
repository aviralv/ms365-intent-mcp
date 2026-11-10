# ms365-intent-mcp

An intent-oriented MCP server for Microsoft 365.

Most Microsoft 365 MCP servers replicate Graph API endpoints as tools — one tool per endpoint, orchestration pushed to the LLM. This server takes a different approach: tools map to user intents, not API structure. One tool call answers one question.

## Tools

| Tool | What it does |
|------|-------------|
| `resolve` | Takes any M365 URL (Teams message, email, file, meeting) and returns its content |
| `my_day` | Your day at a glance: calendar with bodies, mail summary, Teams activity |
| `meeting` | Everything about a meeting: metadata, body, attendees, transcript, chat |
| `compose` | Create email drafts or Teams messages; create, forward, or update calendar events |
| `whats_new` | Unified activity feed across calendar, mail, and Teams for a time window |
| `find` | Search across all M365 data (mail, files, SharePoint, Teams) |
| `people` | Person lookup with recent interaction context |
| `schedule` | Find available meeting times using the scheduling assistant |
| `transcript` | Download a meeting recording's VTT (by URL, meeting name, or item_id+drive_id+site_root), or `list=true` to enumerate recordings newest-first |

### Targeted calendar-event editing

```python
# Synthetic example: resize one occurrence from 13:00–16:00 to 14:00–16:00 CET.
# Use the exact events[i].id from the structured my_day response, not a subject.
compose(payload={
    "type": "event", "mode": "update", "event_id": "synthetic-occurrence-id",
    "start": "2026-11-13T14:00:00", "timezone": "Europe/Berlin",
})
```

- **Partial PATCH:** `event_id` and at least one of `subject`, `start`, `end`, or `location` are required. Omitted fields are not sent to Graph. Null updates, empty subjects, timezone-only requests, and requests with no effective change are rejected. An explicit `location=""` clears the location. Body, attendees, online-meeting settings, `to`, and `comment` are deliberately unsupported in update mode, even if supplied as empty/default values; leave them out. Create and forward contracts remain unchanged.
- **Time safety:** naive timestamps require an explicit valid IANA timezone. Offset-aware timestamps identify an instant; an optional IANA timezone controls interpretation/display, never reinterprets the offset. Without one, aware updates and the response use UTC, not the configured default timezone. Ambiguous/nonexistent naive DST wall clocks are rejected; specify an offset to distinguish a repeated hour. Changed endpoints are sent as UTC `dateTimeTimeZone` pairs to preserve DST-fold instants. Omitted endpoints are not rewritten. The composer fetches the existing event in UTC and checks `end > start` using actual instants, including an omitted endpoint. Unlike create, update has no 12-hour limit, so longer existing events can be edited without truncation.
- **Target and notifications:** updates address `/me/events/{event_id}` only, never search by subject or touch neighboring events. Use an occurrence ID for a recurring meeting; series masters and all-day time edits are rejected. Existing attendees, body (including Teams metadata), recurrence, and other omitted details are preserved. Graph can send meeting-update notifications when editing a meeting with attendees; this is an immediate edit, not a draft. No delete/cancel or scheduling automation is added.
- **Result and errors:** success is `EventUpdated(type="event_updated", event_id, subject, start, end, timezone, web_link, rendered_markdown)` with offset-aware times from Graph's updated response. Human markdown shows the subject, both times, and calendar link, never raw IDs. The existing compose MCP `result` wrapper is retained and its output schema now accepts `ErrorResponse` as well as success responses. Missing `Calendars.ReadWrite`, read/PATCH failures, and validation failures never produce a success response. When Graph supplies an ETag it is sent as `If-Match`; a stale write (412) surfaces as `graph_api_error`, without an unconditional retry. Without an ETag the read/validation/write is not atomic. After an ambiguous transport failure or response-validation failure, re-read the event before retrying: the mutation may already have completed. The existing optional 10-minute idempotency cache stores successful responses only.

### Occurrence-safe meeting lookup

```python
# Synthetic examples; identifiers below are not real Graph IDs.
meeting(payload={"identifier": "Synthetic Weekly Sync", "date": "2026-11-09"})
my_day(payload={"date": "2026-11-09", "include_bodies": True})
# Pass an events[i].id from the structured my_day response:
meeting(payload={"identifier": "synthetic-occurrence-id"})
```

- **Explicit `date`** must be `YYYY-MM-DD` and selects an occurrence whose **start date** is that day in the configured calendar timezone (`default_timezone`, UTC by default). The lookup uses Graph `calendarView` to expand recurring events, with timezone-offset day bounds and a next-midnight boundary that handles DST. Events overlapping from a previous day are not selected. If several subject-substring matches start that day, the earliest is selected; use the event ID to disambiguate further.
- **Date misses** return `ErrorResponse(code="not_found")`, never a different-date occurrence. The MCP output schema accepts either a flat `MeetingDetail` or flat `ErrorResponse`; neither is wrapped in a new `result` field. A valid event ID with a date is checked against that date without subject fallback; a series-master ID with a date is rejected (`validation_error`), so use an occurrence ID from `my_day`. `date` cannot be combined with `"next"`. Unsupported configured timezones are rejected rather than guessed.
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

Requires Python **3.11+** and FastMCP **4.1.6+**. The development lockfile pins FastMCP 4.1.6 for reproducible validation; normal tool installs resolve compatible dependencies from the declared minimum.

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
