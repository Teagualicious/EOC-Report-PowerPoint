# STATUS

> The single source of truth for project state. Every session reads this first and updates it last.

## Current phase

Phase 2 — Mapper roadmap Phase 3: MappingModel extraction (complete, released as v1.24.0)

The Spectrum Reach Reporting Ingestion Engine (from the IngestionEngine_FillTracking handoff zip, 2026-07-12 build) lives at the repo root. See `AI_CONTEXT.md` and `documentation/MODEL_HANDOFF.md` before touching application code. Releases so far: v1.23.0 (repository integration), v1.24.0 (MappingModel extraction).

## Done

- Phase 0 repo cleanup: removed committed `tests/__pycache__`, added `.github/workflows/ci.yml` (pytest on push/PR, Python 3.12). (Root `.gitignore` was actually empty until Phase 1 — see below.)
- Phase 1 (2026-07-13):
  - Imported the inherited codebase, restructured to repo standards: tests moved `developer/tests/` → `tests/`, root `requirements.txt` includes `app/requirements.txt` + pytest, pywin32 gained a `sys_platform == "win32"` marker, root `.gitignore` populated (runtime workspace state, input/output data, caches, secrets).
  - Verified every current doc against the code and fixed stale/wrong content (test counts, paths, layout diagrams; added July 12 fill-report/preview-health coverage; removed false `Workbook_Open` claim; corrected API_REFERENCE signatures). Regenerated `PROJECT_MANIFEST.json`. Removed the generated PDF guides.
  - Extended the fill-engine golden suite with 7 edge-case characterization tests (image path fallbacks, corrupt-image isolation, image-over-text precedence, date `format_details`, mixed placeholders, no-text-frame no-op).
  - Released as v1.23.0 (PRs #3/#4, Release workflow added).
- Phase 2 (2026-07-13) — Mapper roadmap Phase 3, MappingModel extraction:
  - New `app/mapper/mapping_model.py`: `MappingModel` owns all template-mapping state (assignments, images, skip, per-metric format prefs). Pure Python — fully covered by the automated suite.
  - `mapper_window` / `slide_view` / `format_popup` route every mutation through the model; the wizard subscribes and re-renders the shape panel + live COM preview from model state on each change. `wizard.mapping` and `wizard._metric_formats`/`_metric_format_details` are now read-only views onto the model.
  - Persisted schema unchanged (incl. legacy shape-level single assignments); assignment semantics preserved verbatim (update-in-place, replace-all confirmation, format propagation).
  - 15 new model unit tests; 236 tests pass; `python -m compileall -q app tests` clean. All 221 pre-refactor tests pass unchanged (the golden-suite acceptance gate).

## Next up

1. Windows/Office acceptance pass — Excel VBA injection, PowerPoint COM live preview, fill-summary dialogs, `fill_history.jsonl`, and now the refactored mapper (assign/skip/clear/format flows with live preview running) — per `documentation/TESTING_AND_RELEASE.md` and the drill in `documentation/reviews/MAPPER_RELIABILITY_ROADMAP_2026-07-12.md`. Nothing COM-related may be declared verified until this passes.
2. Mapper roadmap Phase 4 — stable shape identity (`shape.Id` with positional fallback); deliberately updates `test_scan_and_fill_agree_on_shape_identity`; everything else must pass unchanged.
3. Mapper roadmap Phase 5 — small fixes from the July 11 review (template-preview debounce, client-wizard drag-select dead code, zoomed-vs-fit_window, all-caps case-forcing decision, skip-discards-assignments decision — see Noticed).

## Decisions log

<!-- Date — decision — why. Keeps future sessions from re-litigating settled questions. -->
- YYYY-MM-DD — Repo created from template —
- 2026-07-12 — CI targets Python 3.12 only — matches current toolchain; matrix can be added later if multi-version support is needed.
- 2026-07-13 — Inherited app merged at repo root (not a subfolder); tests live in root `tests/` — matches CLAUDE.md standards; `config/paths.py` derives all paths from the app location, so the app is unaffected. Docs updated accordingly.
- 2026-07-13 — Root `requirements.txt` is the dev/CI install and includes `app/requirements.txt` — one source of truth for versions; the Windows launcher keeps installing from `app/requirements.txt` unchanged.
- 2026-07-13 — Generated PDF guides removed; markdown is canonical — binary docs can't be reviewed in PRs and were already stale.
- 2026-07-13 — Layout test asserts `PROJECT_ROOT == dirname(APP_DIR)` instead of a hardcoded `IngestionEngine` folder name — the contract is the structure, not the folder's name.
- 2026-07-13 — Each PR gets a release (CLAUDE.md workflow rule 5), created by the **Release** workflow (`.github/workflows/release.yml`) with the top CHANGELOG entry as notes; version continues the inherited line (v1.22 = frozen NYC demo build). v1.23.0 released this way. Remote sessions cannot push tags or call the release API — the workflow is the supported path.
- 2026-07-13 — MappingModel keeps the persisted mapping schema byte-compatible (legacy shape-level single assignments included) — old user mappings and `engine.pptx_fill` must keep working without migration.
- 2026-07-13 — Chart/table live-preview updates stay outside the model — they use transient query-builder data that is not mapping state and re-adding them is not idempotent.
- 2026-07-13 — v1.24.0 released (PR #5, MappingModel extraction) via the Release workflow.

## Noticed (not yet acted on)

<!-- Problems spotted mid-task but out of scope. Harvest these periodically. -->
- Launcher and app docs target Python 3.10+, while repo standard is 3.11+ (CI 3.12). Align when there's a reason to touch the launcher.
- `AI_CONTEXT.md` §13/§17 still describe delivering a "clean archive" — stale process language now that the project is a git repo with CI.
- `tests/test_parsers.py::test_real_input_files` scans root `input/` for optional HTML samples — always empty in the repo/CI, so it's effectively a no-op there.
- Text assignments mapped onto shapes without a text frame are silently unreported by `FillReport` (locked as a characterization test); decide whether to surface it.
- Two `tests/test_ui_helpers.py` DPI tests import `ui.utils` and therefore need a Python built with tkinter (GitHub Actions setup-python and standard Windows installs have it; minimal Linux builds may not).
- Toggling a shape's Skip checkbox replaces its whole mapping and silently discards existing assignments (pre-model behavior, deliberately preserved and locked by `test_set_skip_replaces_shape_mapping`). Decide in mapper Phase 5 whether skip should preserve assignments.
- `PPTXWizard.image_paths` is an unused leftover attribute.

## How to run

```
pip install -r requirements.txt
pytest
python app/main.py   # launches the Tkinter app (Windows users: Start Ingestion Engine.bat)
```
