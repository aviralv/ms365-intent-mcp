# Changelog

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
