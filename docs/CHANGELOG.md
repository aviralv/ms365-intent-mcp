# Changelog

## 1.14.0 — 2026-11-10

### Added
- `compose` event `mode="update"` by exact Graph event/occurrence ID, with partial `subject`, `start`, `end`, and `location` updates. Omitted fields (including attendees, body, recurrence, online-meeting settings, and an unchanged time endpoint) are not sent in the PATCH. Series masters and all-day time edits are rejected; there is no delete/cancel operation.
- Strict timezone/DST and actual-instant ordering validation against current UTC event timing, UTC serialization of changed endpoints, no-op/null/default guards, conditional `If-Match` when Graph returns an ETag, and an offset-aware `EventUpdated` response with human-readable subject/times/link. Existing create/forward behavior and compose's MCP `result` wrapper are preserved; the advertised output union additionally accepts existing `ErrorResponse` outcomes.
- Mocked schema, dispatch, Graph PATCH, preservation, CET/CEST/DST, partial timing, failure, and MCP-client contract regressions. No live calendar mutations were used in development.

## 1.13.1 — 2026-11-09

### Fixed
- `meeting` now advertises both `MeetingDetail` and `ErrorResponse` in its MCP output schema, so explicit-date misses, ID/date mismatches and Graph errors reach consumers instead of failing output validation. Both success and error responses retain their flat shape; the explicit object-root union avoids FastMCP's inferred `result` wrapper.
- Added synthetic MCP-client integration regressions that validate actual tool results against the advertised output schema, rather than only exercising the implementation directly.

### Changed
- Raised the FastMCP minimum to `4.1.6` (current stable) and refreshed the lockfile from `3.4.2` to `4.1.6`, including required framework dependency updates. Python `>=3.11` remains supported; the lockfile is retained for reproducible test/install environments.

## 1.13.0 — 2026-11-09

### Added
- Optional `meeting(payload={identifier, date})` occurrence start date (`YYYY-MM-DD`, configured calendar timezone). Explicit-date subject searches use day-scoped Graph `calendarView`, validate the selected occurrence, and never fall back to another date. Event IDs remain supported; date plus an occurrence ID validates its date, while series-master IDs and date plus `next` are rejected. Omitted-date lookup behavior is unchanged.
- `my_day.events[].id` exposes each Graph event/occurrence ID for roundtripping into `meeting`, without adding IDs to rendered markdown or changing existing event fields.

### Fixed
- Explicit-date recurring meeting lookup no longer selects the earliest subject match in a ±7-day window. Date misses return `not_found`; existing recording occurrence-date matching and stale warnings are preserved.

## 1.5.0 — 2026-07-17

### Added
- **`people()` chat-membership fallback (#44).** When `/me/people` and `/me/contacts` both miss — as they do for first-name-only or SAP-internal queries like `people(query='Yevhen')` — `people()` now enumerates the user's Teams chats and synthesizes matched people from chat members, so a live 1:1 chat resolves the person and surfaces its `chat_url`. Self-exclusion and dedup key on the AAD object id (`member.userId`), falling back to email then displayName. The `/me/chats` list is fetched once and reused for both the fallback and `recent_chat` enrichment.

## 1.0.0 — 2026-07-08

### Removed
- **Legacy flat-kwargs tool surface.** Every tool now takes `payload={...}` with Pydantic validation and returns a typed response. See v0.8.0 CHANGELOG for the shape.
- Server-side deprecation warnings (no more legacy path to warn about).

### Renamed
- `my_day_v1` → `my_day`
- `meeting_v1` → `meeting`
- `compose_v1` → `compose`
- `schedule_v1` → `schedule`
- `people_v1` → `people`
- `whats_new_v1` → `whats_new`
- `find_v1` → `find`
- `resolve_v1` → `resolve`

### Migration for consumers
Consumers on v0.8.0 that call `<tool>_v1` should update to the unsuffixed name. Behavior is identical — same payload shape, same typed response shape. Only the tool name changes.

Consumers still on the v0.7.x flat-kwargs surface: this is a breaking change. Migrate to the v0.8.0 payload shape first (see v0.8.0 CHANGELOG), then rename `_v1` → canonical.

## 0.8.0 — 2026-07-08

### Added
- **`intent/` package** — new Pydantic-based tool surface. Every tool exposes a `_v1` variant with `payload={...}` calling convention and typed Pydantic response models.
  - Pydantic discriminated union on `compose_v1` (3 variants: `email`, `event`, `teams_message`)
  - Every response inherits from `BaseResponse` with a `type` Literal discriminator
  - Uniform `ErrorResponse` envelope (`code`, `message`, `retryable`)
  - `rendered_markdown` field on every response — machine consumers read structured fields, human consumers read the markdown
  - Idempotency keys on write variants (10-min TTL, in-memory cache)
- **Golden JSON Schema snapshots** in `tests/intent/snapshots/schemas/` — catches silent schema drift from Pydantic upgrades or accidental field changes.
- **Live-verify scripts** in `scripts/verify_*_v1.py` — one per tool, repeatable smoke tests against live Graph.

### Changed
- Composers now return `tuple[dict, str]` instead of `str`. Legacy tools unpack `_, markdown` to preserve their existing behavior.
- Legacy tools now log a `_logger.warning("legacy tool '<name>' called — migrate to '<name>_v1'")` on every call. Warning is server-side only; not surfaced to LLM responses.

### Migration
- Existing consumers keep working — legacy tool names still resolve to the same behavior as v0.7.x.
- New consumers should use `_v1` names. In v1.0.0, `_v1` variants will be renamed to their canonical names and legacy variants removed.
- Deprecation timeline: 7 consecutive days of zero deprecation warnings in server logs → v1.0.0 cutover.

### Fixed
- (nothing user-facing — this release is entirely additive on the input surface)
