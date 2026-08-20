# STATUS

> Single source of truth for Deck Engine project state. Read first; update last.

## Current phase

**Stage 1 complete (2026-08-20) — ready to enter Stage 2.**

Stage 0 (fork surgery and architecture harness) is done and merged (PR #1). A planning
session on 2026-08-20 locked the product direction with the owner, produced
`documentation/HANDOFF.md` (context, architecture map, phase plan, decisions DEC-1…DEC-8),
amended `documentation/DECK_ENGINE_BUILDOUT.md` (Stage 5 is now a localhost web dashboard),
and added binding `documentation/DOCUMENTATION_STANDARDS.md` and
`documentation/TESTING_STANDARDS.md`.

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
- [x] Stage 0 branch published as draft PR #1.
- [x] Stage 1 ingestion core implemented on the local `codex/stage-1-ingestion`
  branch: CSV/XLSX/XLSM/HTML inspection, stable structure fingerprints,
  schema-versioned import profiles, explicit profile-required refusal, v0
  campaign-dictionary passthrough, reconciliation output, and deterministic
  synthetic fixtures.
- [x] Stage 1 tests cover known-profile replay, column-order invariance, structural
  changes, malformed inputs, all supported formats, dictionary notes, and the
  50,000-row performance budget.

## Planning pass and Stage 1 implementation 2026-08-20 — what changed and checks run

- Added `documentation/HANDOFF.md`; amended `DECK_ENGINE_BUILDOUT.md` (Stage 5 →
  localhost dashboard; Stage 4 + interim branded template; Stage 6 launcher); added
  `DOCUMENTATION_STANDARDS.md`, `TESTING_STANDARDS.md`; updated `documentation/README.md`
  and `CLAUDE.md` reading order. Stage 1 then added
  `app/engine/ingestion.py`, `app/engine/synthetic_fixtures.py`, the v0 dictionary
  implementation, parser error checks, and contract tests. The Stage 0 architecture
  label and logging-backup documentation drift were corrected on the next code touch.
- Checks: `python -m compileall -q app tests` clean; `python -m pytest -q`:
  **322 passed**. The bundled runtime includes tkinter, so the prior container-only
  Tk import gap did not occur in this run.
- The plan was reviewed by a three-lens adversarial panel (architecture/feasibility,
  cold-start pickup, process/standards) before publishing; the report, ratings, and issue
  dispositions live in `documentation/reviews/HANDOFF_REVIEW_2026-08-20.md` (committed in
  the same change). All blocker and major findings were fixed in this pass.

## Recorded findings (out of scope for this pass — do not fix silently)

- **`python -m app` bootstrap bug:** crashes with `ModuleNotFoundError: No module named
  'config'` (exit 1) instead of `app/main.py`'s intended exit 2 — `app/__main__.py` lacks
  the `sys.path` bootstrap `app/cli.py` has. Fix scheduled with the Stage 5 entry-point
  rework (or any earlier code-touching pass).
- **`documentation/upstream/` contains no EOC requirements document** despite the old
  index claiming so (index corrected this pass); if a sanitized requirements/defect
  review exists outside the repo, the owner may supply it for `documentation/`.

## Next

1. **Stage 2 entry condition:** read `documentation/HANDOFF.md` §7 Phase 2 and
   `documentation/DECK_ENGINE_BUILDOUT.md` Stage 2; implement the literal,
   versioned staging workbook writer/reader from the parsed Stage 1 contract.
2. Preserve the Stage 1 profile/refusal behavior and validate against a sanitized
   real export only when one is supplied; until then, real-file contact remains
   RSK-2 (owner confirmed synthetic-first, DEC-4).
3. Open questions Q1–Q5 in `HANDOFF.md` §10 have named decide-at stages — answer
   them there, not before.

## Decisions log

- **2026-08-20 — DEC-1 Analyst UI is a localhost web dashboard** (owner); the Tk window
  plan is dropped. **DEC-2 Metrics edited in desktop Excel** via the dashboard's
  Open-in-Excel; no in-browser editing surface (owner). **DEC-3 Per-page content
  checkboxes** wireframed behind a default-off flag until the real template arrives
  (owner). **DEC-4 Synthetic-first continues** (owner). **DEC-5 Template-first IR is the
  production build carrier**, classic path kept for regression/dev — confirm at Stage 4
  entry. **DEC-6 Interim Spectrum-branded default template** committed and replaceable.
  **DEC-7 Dashboard server stdlib-first**, localhost + token. **DEC-8 Chart/table
  payloads are literal workbook content.** Full text: `documentation/HANDOFF.md` §3–4.
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
