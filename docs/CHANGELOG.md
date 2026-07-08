# Changelog

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
