# Changelog

## July 13, 2026 - Repository integration

- Merged the inherited IngestionEngine codebase into the Jughead-Data-Engine
  GitHub repository; the project now lives at the repository root instead of
  an `IngestionEngine/` folder.
- Moved the pytest suite from `developer/tests/` to root `tests/` per repo
  standards; test command is now `python -m pytest tests -q` from the root.
- Added a root `requirements.txt` (runtime deps from `app/requirements.txt`
  plus pytest) so `pip install -r requirements.txt && pytest` works anywhere;
  the Windows launcher still installs from `app/requirements.txt`. pywin32
  gained a `sys_platform == "win32"` marker so Linux CI installs cleanly.
- GitHub Actions CI now runs the suite on every push and pull request.
- Removed the generated USER_GUIDE.pdf and TECHNICAL_GUIDE.pdf; the markdown
  guides are canonical.
- Refreshed all current documentation against the code (test counts, paths,
  layout diagrams, fill-report/preview-health coverage) and regenerated
  PROJECT_MANIFEST.json.
- Extended the fill-engine golden suite with 7 characterization tests:
  workspace-relative and `image_path_abs` image fallbacks, corrupt-image
  error isolation, image-over-text assignment precedence, date-styled
  `format_details`, mixed matched/unmatched placeholders on one shape, and
  the silent no-op for text assignments on shapes without a text frame.
- 221 automated tests pass (212 inherited + 2 repository smoke tests +
  7 new golden tests).

## July 12, 2026 - Mapper success/failure tracking + golden-file safety net

- Added a golden-file characterization suite for the previously untested fill
  path: 19 tests on engine/pptx_fill and 8 on engine/pptx_mapper
  (save/load, scanning, scan-fill identity contract).
- Added engine/fill_report.py: every fill now produces a FillReport (filled,
  images, skipped, missing metrics, unmatched placeholders, missing images,
  failed queries) shown in the Save & Fill / Auto-Fill dialogs and appended
  to workspace/logs/fill_history.jsonl.
- Unmatched replace_text placeholders — previously a silent no-op — are now
  reported by name.
- Added health tracking to all 8 public COM methods in engine/pptx_live:
  after 3 consecutive PowerPoint errors the live preview disables itself,
  notifies the user once, and Save & Fill falls back to the built-in
  python-pptx engine automatically.
- fill_template() signature unchanged (delegates to fill_template_report),
  so existing callers and mappings are unaffected.
- 212 automated tests pass (was 171). Windows COM verification checklist in
  documentation/reviews/MAPPER_RELIABILITY_ROADMAP_2026-07-12.md.

## July 10, 2026 - Windows laptop UI corrections

- Used real Windows laptop photos to correct high-DPI clipping that was not visible in the Linux smoke test.
- Made window geometry scale with Windows display scaling while remaining inside the usable desktop area.
- Removed the fixed header height so subtitles and workflow steps no longer clip.
- Shortened and widened the Settings tabs and added useful empty states.
- Rebuilt Platform Setup so its Save/Cancel bar stays at the bottom instead of overlapping table headers.
- Made platform-mapping columns resize with the window and expanded the inner canvas to the available width.
- Rebuilt the client-assignment list so campaign rows use the full screen and Next/Back remain in a dedicated bottom bar.
- Made PowerPoint template previews load asynchronously with a slide-text fallback, so the preview no longer remains blank while PowerPoint starts.
- Ensured the template selector always reserves room for Auto-Fill, New Template, and Cancel buttons.
- Added DPI sizing and template-preview regression tests; all 171 automated tests pass.

## July 10, 2026 - AI context handoff guide

- Added root `AI_CONTEXT.md` as a concise operating guide for future AI assistants.
- Documented product priorities, repository boundaries, business invariants, threading rules, compatibility expectations, high-risk files, common failure modes, technical debt, and definition of done.
- Cross-linked the AI guide from the root README, documentation index, model handoff, and project manifest.
- Revalidated the full automated test and compile checks after documentation-only changes.

## July 10, 2026 - Root input/output navigation update

- Moved the user-facing `input/` and `output/` folders from `workspace/` to the main project folder.
- Kept settings, templates, mappings, and logs under `workspace/`.
- Added non-destructive migration from `workspace/input/`, `workspace/output/`, and the older `input_files/` location.
- Automatically translates the previous default `workspace/output/` setting to the new root `output/` folder while preserving custom output paths.
- Updated the user, technical, architecture, testing, API, and model-handoff documentation.
- Confirmed all 167 automated tests pass after the change.

## July 10, 2026 - Documentation and structure handoff build

- Reorganized the project into four clear areas: `app/`, `workspace/`, `documentation/`, and `developer/`.
- Replaced the ambiguous launcher with `Start Ingestion Engine.bat`.
- Moved code and static resources under `app/`.
- Moved settings, templates, mappings, logs, inputs, and output into a dedicated user-data structure (input/output were subsequently moved to the root in the navigation update above).
- Moved tests, PyInstaller configuration, and portable-build tooling under `developer/`.
- Added non-destructive migration from the old root-level data layout.
- Added compatibility resolution for pre-workspace PowerPoint image paths.
- Rewrote current user, technical, architecture, API, testing, and handoff documentation.
- Archived historical state and refactor documents so they are no longer mistaken for current implementation guidance.
- Updated PDF user and technical guides.
- Confirmed all 165 automated tests passed at the time of the initial restructuring.

## July 10, 2026 - Code review, optimization, and UI refresh

- Improved large XLSX import time by about 30% and reduced peak memory in the synthetic benchmark.
- Improved CSV time and memory use.
- Prevented repeat exports from doubling matching rows.
- Chose the best metric source per campaign instead of globally.
- Averaged rate metrics instead of summing them.
- Preserved acronyms such as CTR, CPM, and CPC.
- Fixed UTF-8 BOM handling.
- Removed direct legacy XLS selection and added conversion guidance.
- Fixed client selections disappearing after filtering.
- Added safe, collision-resistant path naming.
- Hardened template ZIP import/export.
- Moved long-running parsing and export work off the Tkinter main thread.
- Refreshed the main workflow and review UI.

See `reviews/CODE_REVIEW_2026-07-10.md` for the detailed review.
