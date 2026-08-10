# STATUS

> Single source of truth for Deck Engine project state. Read first; update last.

## Current phase

**Fork Stage 0 — fork surgery and architecture harness.**

The fork starts from the exact upstream Jughead Data Engine v1.36.0 commit:

- Repository: `Teagualicious/Jughead-Data-Engine`
- Commit: `386e8c1f7a553c7aff40c5445adbc451df97be46`
- Baseline verification on 2026-08-10: **367 tests passed** on the untouched source archive.

## Stage 0 objective

Produce a self-contained survivor codebase before implementing new product behavior:

- strip features listed in Appendix A of `documentation/DECK_ENGINE_BUILDOUT.md`;
- preserve the parser/KPI/mapping/fill cores and entire golden fill suite;
- establish stable workflow/module boundaries;
- enforce path, sanitizer, and headless-engine architecture laws;
- leave the analyst UI intentionally unavailable until Stage 5.

## Completed this stage

- [x] Recorded untouched upstream baseline and test count.
- [x] Removed MCP, VBA/search workbook, COM live preview, and obsolete analyst windows.
- [x] Archived upstream documentation and added the authoritative fork plan.
- [x] Added `workspace/staging`, `workspace/dictionary`, and `output/_quarantine` paths.
- [x] Added stable Stage 1–4 workflow and module scaffolds.
- [x] Moved mapper launch to `developer/run_mapper.bat` / `python -m app.mapper`.
- [x] Removed `pywin32` and `tkcalendar` from runtime requirements.
- [x] Survivor suite and architecture tests green: **313 passed**.
- [ ] Stage 0 branch published as a draft PR.

## Next

1. Publish the Stage 0 branch as a draft PR and verify remote CI.
2. Stage 1: implement single-dump ingestion, stable fingerprints/import profiles, and the v0 campaign-dictionary passthrough using synthetic fixtures.
3. Validate Stage 1 against a sanitized real export when one is supplied; until then, record real-file contact as RSK-2.

## Decisions log

- **2026-08-10 — Upstream history retained.** The target repository already points at the exact v1.36.0 donor commit, so the fork branch is based directly on it rather than re-importing code as an unrelated root.
- **2026-08-10 — Development VERSIONs never release.** `0.1.0-dev` and later `*-dev` versions are CI/test artifacts; the release workflow exits before publishing them.
- **2026-08-10 — Synthetic-first Stage 1.** No sanitized real dump was attached. The plan explicitly permits the deterministic synthetic factory while naming first contact with real data as a risk.
- **2026-08-10 — Mapper is static and developer-only.** The COM live-preview module is removed; mapping and fills remain available through python-pptx.

## Non-negotiable project rules

- No real client data, campaign data, credentials, or company-internal exports in the repository.
- Tests are the completion gate; new behavior receives a new test.
- UI and CLI remain thin shells over `engine.workflow`.
- The staging workbook becomes the only source of truth for deck values in Stage 2/4.
- Never claim Windows/Office behavior verified without an explicit Windows acceptance pass.
