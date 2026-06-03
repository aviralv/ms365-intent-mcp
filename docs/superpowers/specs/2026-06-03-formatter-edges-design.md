# Formatter Edge Cases Design

**Date:** 2026-06-03
**Branch:** `fix/formatter-edges`
**Scope:** Two surgical formatter fixes for edge cases flagged during PR #10 review.

---

## Goal

Land two independent fixes on one branch:

1. **Cross-date event detail rendering** — `format_event_detail_markdown` currently renders the event end as time-only (e.g. `14:30 UTC`). For cross-midnight events this loses the date and produces ambiguous output. Render the end as full date+time, always, matching the start.
2. **Null-`ts` call entry artifact** — `_format_chat_entry`'s call branch produces `"  UTC"` (double space + orphan UTC) when an entry has a missing or empty `ts`. The message and event branches handle this cleanly via `ts_with_tz`. Make the call branch consistent.

Both are pre-existing latent issues, surfaced by reviewers during PR #10. Neither has been observed in production. Fixing now is cheap (~5 lines each) and prevents future traps.

## Why both on one branch

Both touch `formatters.py` + `tests/test_formatters.py` only. Both are surgical (~5 production lines each, ~30 test lines total). Reviewing one PR is cheaper than two. Independently revertible commits.

## What this is NOT

- **Not a fix for issue #9** (`find(type='message')` permission rejection). Two separate experiments confirmed the underlying cause is `ChannelMessage.Read.All` admin-consent gate in the SAP tenant — not something we can fix in this codebase. Issue #9 stays open for separate follow-up: either ask SAP IT (slim odds), build an enumerate-then-filter fallback for personal/DM chats (~3-4 hours, separate plan), or document the limitation and route around with `resolve(<chat URL>)`.
- **Not a scope-list change.** PR #10 already shipped the timezone fix; the original plan to bundle a `Chat.Read` scope addition with these formatter edges has been dropped after the experiment proved the scope addition pointless.
- **Not a re-architecture of formatters.** The two helpers added in PR #10 stay as they are. This branch just calls `_format_event_datetime` once more.

---

## Background context

### Pre-existing reviewer-flagged issues (latents)

During the PR #10 reviews, two reviewers independently flagged code paths with edge-case wrong behavior:

- `format_event_detail_markdown` Task 5 review: *"a late-night event (23:30–00:15 UTC) renders 00:15 with no date. A consumer reading this has no way to know the end is the next day."*
- `_format_chat_entry` Task 11 review: *"the call branch doesn't use ts_with_tz at all — when ts is null, output is '  UTC' — two orphaned spaces plus UTC — where message/event branches would silently produce empty parens."*

Both are pre-existing bugs visible only on edge inputs Graph doesn't currently send. Defensive fixes only — but cheap.

---

## Architecture

Two commits, in this order:

```
1. fix(formatters): render full date+time on event detail end
2. fix(formatters): render empty parens for null-ts call entry
```

Each commit is independently revertible. No cross-dependency.

### Files touched

| File | What changes |
|---|---|
| `src/ms365_intent_mcp/formatters.py` | (#1) Replace inline end-time block in `format_event_detail_markdown` with a second `_format_event_datetime(end)` call. (#2) Add `if not raw_ts:` guard at top of `_format_chat_entry` call branch. |
| `tests/test_formatters.py` | Update existing assertion in `TestFormatEventDetailMarkdown.test_includes_all_fields`. Add `test_cross_date_event_renders_both_dates`. Add `test_call_with_null_ts_renders_empty_parens`. |
| `pyproject.toml` | Bump 0.4.1 → 0.4.2 |

No other files modified. No composer changes. No tool API changes. No scope changes.

---

## Detailed changes

### 1. Cross-date event detail end rendering

**`src/ms365_intent_mcp/formatters.py`** `format_event_detail_markdown` lines 107-113 currently:

```python
lines = [f"## {subject}"]
start_fmt = _format_event_datetime(start)
end_dt = end.get("dateTime", "")
end_tz = end.get("timeZone") or start.get("timeZone") or ""
end_hm = end_dt[11:16] if len(end_dt) >= 16 else end_dt
end_fmt = f"{end_hm} {end_tz}".strip()
lines.append(f"**When:** {start_fmt} → {end_fmt}")
```

Becomes:

```python
lines = [f"## {subject}"]
start_fmt = _format_event_datetime(start)
end_fmt = _format_event_datetime(end)
lines.append(f"**When:** {start_fmt} → {end_fmt}")
```

Net: -4 lines.

**Behavior change:**

| Case | Before | After |
|---|---|---|
| Same-day | `**When:** 2026-05-15T14:00 UTC → 14:30 UTC` | `**When:** 2026-05-15T14:00 UTC → 2026-05-15T14:30 UTC` |
| Cross-day | `**When:** 2026-05-15T23:30 UTC → 00:15 UTC` (date lost) | `**When:** 2026-05-15T23:30 UTC → 2026-05-16T00:15 UTC` |

The same-day case becomes slightly more verbose. Trade-off: consistent rendering across single-day and cross-day events; consumers always see the date, no special-case rules to remember.

### 2. Null-`ts` call entry empty parens

**`src/ms365_intent_mcp/formatters.py`** `_format_chat_entry` call branch (current):

```python
if kind == "call":
    raw_end = (entry.get("end_ts") or "")[:16]
    start_date, start_time = raw_ts[:10], raw_ts[11:16]
    end_date, end_time = raw_end[:10], raw_end[11:16]
    if not raw_end or raw_end == raw_ts:
        time_range = f"{start_date} {start_time} UTC"
    elif start_date == end_date:
        time_range = f"{start_date} {start_time} UTC–{end_time} UTC"
    else:
        time_range = f"{start_date} {start_time} UTC → {end_date} {end_time} UTC"
```

Becomes (one new branch at the top):

```python
if kind == "call":
    raw_end = (entry.get("end_ts") or "")[:16]
    start_date, start_time = raw_ts[:10], raw_ts[11:16]
    end_date, end_time = raw_end[:10], raw_end[11:16]
    if not raw_ts:
        time_range = ""
    elif not raw_end or raw_end == raw_ts:
        time_range = f"{start_date} {start_time} UTC"
    elif start_date == end_date:
        time_range = f"{start_date} {start_time} UTC–{end_time} UTC"
    else:
        time_range = f"{start_date} {start_time} UTC → {end_date} {end_time} UTC"
```

Net: +2 lines.

**Behavior change:**

| Case | Before | After |
|---|---|---|
| `ts = ""` or missing, single-event call | `📞 **Call** (  UTC)` | `📞 **Call** ()` |
| All other cases | unchanged | unchanged |

This matches how the message and event branches already handle null `ts` (they use `ts_with_tz` which returns `""` for empty input, producing `()`).

---

## Testing

### Unit tests

Three test additions/changes in `tests/test_formatters.py`:

1. **Modify** `TestFormatEventDetailMarkdown.test_includes_all_fields` — change the `"14:30 UTC"` assertion to `"2026-05-15T14:30 UTC"` to reflect new full-date end rendering.
2. **Add** `TestFormatEventDetailMarkdown.test_cross_date_event_renders_both_dates` — fixture with start `2026-05-15T23:30:00 UTC` and end `2026-05-16T00:15:00 UTC`; assert both dates appear.
3. **Add** `TestFormatChatEntry.test_call_with_null_ts_renders_empty_parens` — fixture with `ts: ""`, kind: "call"; assert no `"  UTC"` artifact, output contains empty `()`.

Existing tests should continue to pass. Full suite count goes from 75 → 77.

### Manual verification

Not required. These are pure formatter changes covered by unit tests. A session-level smoke check via `meeting()` is sufficient if anyone wants to eyeball the new output.

---

## Out of scope

These came up during the brainstorm and are explicitly NOT part of this branch:

- **Issue #9 fixes** — confirmed not solvable from inside this codebase without admin consent for `ChannelMessage.Read.All`. Separate plan needed for Path B fallback (enumerate-then-filter on `/me/chats`) or Path A pre-flight friendly error.
- **Issue #12 (token-write atomicity)** — separate branch.
- **Issue #8 (recording metadata exposure)** — separate plan, larger feature.
- **Other timestamp paths in formatters/composers** — only the call sites enumerated above are touched. No general sweep.

---

## Verification steps (for the implementing agent)

1. Confirm on branch `fix/formatter-edges` (already created from main).
2. Implement commit 1 (cross-date end). Verify:
   - `format_event_detail_markdown` no longer has the inline `end_hm`/`end_tz` code.
   - The modified existing test passes; the new `test_cross_date_event_renders_both_dates` passes.
   - Test count goes from 75 → 76.
3. Implement commit 2 (null-ts call). Verify:
   - `_format_chat_entry` call branch has the new `if not raw_ts: time_range = ""` line.
   - The new `test_call_with_null_ts_renders_empty_parens` passes.
   - Test count goes from 76 → 77.
4. Bump `pyproject.toml` to 0.4.2. Commit.
5. `uv run pytest -v` — full suite passes (290 → 292).
6. Push branch. Open PR. Wait for user approval. Merge per `.claude/rules/DEPLOYMENT.md`.

---

## Context preserved for follow-up work

The `Chat.Read` scope experiment ran. Result: confirmed not a fix. Both `Chat.Read` and `Chat.ReadWrite` literally present in the JWT scp claim, Graph Search API still rejected `chatMessage` queries with the same error. Subsequent test confirmed the actual gate: `ChannelMessage.Read.All` admin consent, blocked at the standard Graph permission level (not a tenant-specific policy that could be worked around with a different client ID).

For #9 follow-up, three viable directions:

- **Path A** (~1 hour): pre-flight scope check in `composers/find.py` — detect missing `ChannelMessage.Read.All`, surface a friendlier error explaining the limitation and pointing at `resolve(<chat URL>)` as the workaround for known chats.
- **Path B** (~3-4 hours): enumerate-then-filter fallback on `/me/chats` for personal/DM chats only. Doesn't cover Teams channel messages (those still need the admin scope). New brainstorm + plan needed.
- **Path C** (Microsoft 365 admin route): file an SAP IT request for `ChannelMessage.Read.All` on the personal app registration. Slim odds.

These are deferred. This branch ships only the formatter edges.
