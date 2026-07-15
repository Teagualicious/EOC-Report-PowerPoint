# Changelog

## July 15, 2026 - Saved-query fidelity, mapper polish, template-first Phase A

- **Saved builder queries now re-resolve to exactly the value shown when
  they were applied.** The resolver previously ignored the Advanced Query
  Builder's campaign/breakdown/value/Top-N selections, so a report
  refilled in a later session could silently diverge from the pivot the
  user previewed. Builder queries now recompute through the same pivot
  engine (moved to `engine/pivot.py`, shared by the builder UI and the
  resolver — and importable headlessly by the CLI/MCP tools).
- **Skip no longer discards assignments.** Toggling a shape's Skip
  checkbox is now a flag on top of its mapping — untoggle and everything
  is back. (Previously an exploratory click silently wiped the shape's
  assignments.)
- Mapper polish: template-selector previews debounce (no more one
  PowerPoint thumbnail export per arrow-key step), dead drag-select code
  removed from the client wizard, and the all-caps case-matching rule is
  now an explicit product decision: inserted values match the deck's
  case ("CLIENT NAME" placeholder → "ACME HOLDING").
- Releases now publish **two download zips**: a standard one
  (`IngestionEngine-<version>.zip`, smaller — installs packages online on
  first launch) and the portable one (`…-portable-win64.zip`, bundles
  every package for offline first-run setup). Release notes explain which
  to pick.
- **Template-first mapper Phase A shipped** (`app/engine/template_ir/`):
  ingest any deck into a JSON template schema + extracted assets, then
  rebuild a pixel-identical deck from verbatim shape XML — images
  re-linked, charts skipped loudly with a report (chart cloning is
  Phase C). The ingest→build→re-ingest round-trip is locked by tests and
  runs entirely in CI (no PowerPoint, no COM). This is the foundation for
  the slot-based mapper rework documented in
  `documentation/proposals/TEMPLATE_FIRST_MAPPER_2026-07-15.md`.
- 306 automated tests pass.

## July 15, 2026 - Portable download + data-hygiene scrub (demo readiness)

- **Every release now includes a portable zip**
  (`IngestionEngine-<version>-portable-win64.zip`): the program plus a
  `wheelhouse/` of Windows packages for every dependency. Download,
  extract, double-click `Start Ingestion Engine.bat` — the first launch
  installs the bundled packages automatically, offline, in under a
  minute. Only Python 3.12 needs to be on the machine. Release notes now
  tell downloaders which zip to grab.
- The launcher understands all distribution shapes and picks
  automatically: a fully portable `app/vendor/` folder (no install at
  all — previously the launcher ignored it and wrongly attempted a pip
  install on offline machines), the bundled wheelhouse (offline
  install), or the internet (as before).
- **Data-hygiene scrub ahead of external demos:** all example and test
  client names are now clearly fictional ("Acme Appliance Co",
  "Acme Motors"), and the handful of real-but-unattributed campaign
  figures that had crept into docs, test docstrings, and the stand-in
  template were genericized. A full audit found no client exports,
  credentials, file paths, or personal information in the repository.
- 297 automated tests pass.

## July 14, 2026 - Multi-line text fills land reliably; repeated report runs stay stable

- **Assigned text no longer vanishes from exported decks.** Two root causes,
  both in the built-in fill engine (the path Auto-Fill Report uses):
  - A replace target selected across lines carries line-break characters
    (paragraph joins, PowerPoint soft breaks) that the text runs being
    searched never contain — the replace silently did nothing while the
    fill report claimed "filled". Multi-line targets are now matched line
    by line: the first target line receives the value, the remaining
    target lines are cleared, and the report reflects the ACTUAL outcome —
    a failed replace is always surfaced as an unmatched placeholder.
  - A multi-line value (e.g. custom text with several lines) was written
    as a literal newline inside the text run, which PowerPoint renders as
    whitespace — it looked right in the live preview but flattened or
    visually vanished in the exported file. Values now produce real
    PowerPoint line breaks with the run's formatting preserved.
  - Bonus: assigning a full-text value to an empty text box used to no-op
    silently; it now writes the value.
- **The program stays healthy across repeated report generations.** The
  live preview released its COM apartment on every cleanup — and cleanup
  runs twice per mapper session (explicit close + garbage collection) — so
  the UI thread's Office automation degraded after a few cycles ("works
  twice, errors the third time"). COM is now initialized/released exactly
  once per preview. The COM template scan no longer leaks an open
  PowerPoint presentation on failure, and template-thumbnail exports are
  serialized so rapid clicking can't pile up PowerPoint processes.
- 297 automated tests pass.

## July 14, 2026 - Dynamic output folder, faster exports, cleaner review KPIs

- **Renaming the project folder no longer resurrects the old folder.**
  Saving Settings used to pin the absolute default output path into
  settings.json; after a rename, exports recreated the old folder name
  (e.g. `Jughead-Data-Engine-1.27.0\output`) and wrote there. The default
  now stays dynamic: it is never persisted, and a stale default-shaped
  path from a previous location falls back to the current project's
  `output` folder on launch. Custom output folders are untouched.
- **The export stage (client selection → review) is faster and shows live
  progress.** A multi-client export now launches ONE hidden Excel process
  for the whole batch instead of one per client (Excel startup was the
  dominant fixed cost of adding the interactive search), and the workbook
  row writer was rewritten around a faster iteration path for large files.
  The status line now reports each stage as it happens ("writing workbook
  + search…", "computing KPIs…") instead of a single static message.
- **Reach and Frequency no longer appear in the review KPIs** — totals or
  per-campaign detail. Campaign aggregates cannot be household-deduplicated,
  so the previous "not deduplicated" labels still invited comparison with
  vendor dashboards. They return when real household-level data/formulas
  are available; the raw values remain in the exported workbook.
- **KPI numbers are formatted by what they are, everywhere.** Counts render
  as whole comma-grouped numbers (no more trailing ".00" from float noise),
  rates get a % sign ("87.65%"), money gets $ — consistently across KPI
  cards, campaign summaries, and detail rows.
- 291 automated tests pass.

## July 14, 2026 - Stable shape identity for the PowerPoint mapper

- Template fills now find each mapped shape by its persistent PowerPoint
  shape id (with a unique-name fallback) instead of its position in the
  slide. Editing a mapped template — adding, deleting, or reordering
  shapes — no longer sends values into the wrong shapes ("SHAPE INDEX
  DRIFT"), the mapper's flagship correctness bug.
- Scans emit the id as `shape_uid`; the mapper stamps it (plus the shape
  name) onto every mapping entry the user touches. Existing saved
  mappings keep working unchanged and upgrade lazily as they're edited —
  no migration, and old app versions ignore the new fields.
- If a mapped shape was deleted from the template, the fill skips it and
  reports it in the fill summary ("Mapped shapes no longer in the
  template") instead of silently writing into whatever shape inherited
  the old position. The live COM preview retargets by id the same way and
  no longer "warns but writes anyway" on drift.
- One shared pure resolver (`app/engine/shape_identity.py`) drives both
  the built-in python-pptx fill and the Windows live preview; drift
  scenarios (reorder, insert, delete) are locked by end-to-end regression
  tests. 279 automated tests pass.
- Needs the Windows/Office acceptance drill (id parity between python-pptx
  and COM scans, live drift drill) before COM paths are declared verified.

## July 13, 2026 - Typed-pivot search contract + gated KPI cards

- The Excel search now behaves as a strict typed pivot table: result
  columns follow the EXACT typed order — dimensions and metrics
  interleaved as written ("Impressions, Campaign, Client" renders in that
  order). Previously dimensions always rendered before metrics.
- KPI summary cards above the results render only with a reason: for
  metrics the user explicitly typed (never the auto-filled default set,
  which produced unrequested boxes and meaningless roll-ups) and only
  when the table has more than one row (a single row is already its own
  total). Cards keep typed order and the sum/avg/rate math.
- Unrecognized terms still show the "Ignored (no match)" notice; when no
  metric is typed (or none matches), the default metric set appends AFTER
  the typed columns instead of reshuffling them.
- The VBA/dashboard contract (row constants, typed-order rule, KPI card
  gate) is locked by static tests.
- 262 automated tests pass.

## July 13, 2026 - AI-native interfaces: workflow service, terminal CLI, MCP server

- New `app/engine/workflow.py`: the complete reporting workflow (parse
  exports → discover campaigns → client dataset → KPIs → Excel workbook →
  PowerPoint fill → metric queries) as one headless, JSON-friendly service —
  the same engine code the desktop app runs, now callable without the UI.
- New `app/cli.py`: terminal interface designed for AI agents and
  automation. Subcommands `platforms`, `templates`, `scan`, `campaigns`,
  `kpis`, `export`, `fill`, `query`; every command prints a single JSON
  object with `ok`/`data`/`error` and a meaningful exit code. No prompts.
- New `app/mcp_server.py`: local MCP server exposing eight tools
  (list_platforms, list_templates, scan_export, list_campaigns, get_kpis,
  query_metric, export_workbook, fill_template) to Claude Desktop / Claude
  Code over stdio. Campaign data never leaves the machine; the `mcp`
  package is required only where the server runs (not in the app's runtime
  requirements). `INGESTION_MCP_READ_ONLY=1` disables the writing tools.
- New `documentation/AI_INTEGRATION.md`: setup, tool reference, example
  prompts, and governance notes.
- Template mapping remains an interactive task; AI interfaces consume
  saved mappings and keep the usual graceful Office degradation. Every
  AI-triggered fill is logged to `fill_history.jsonl` like any other.
- 11 new tests (workflow service end-to-end, CLI contract incl. JSON
  number types and error envelopes, thin-shell enforcement for CLI/MCP).
  260 automated tests pass.

## July 13, 2026 - Windows debugging batch: wheel/layout fixes, query UX, UI performance

- The mouse wheel no longer changes a closed dropdown's value anywhere in
  the app. Hovering a role/platform combobox while scrolling a pane was
  silently corrupting selections (Platform Setup roles shifted one entry);
  the wheel now always scrolls the surrounding pane instead. Open dropdown
  lists still scroll normally.
- Platform Setup rows now share one grid per sheet, so the Role and Sample
  columns align exactly instead of staggering when sample text is long.
- The client-assignment window no longer maximizes over the Windows
  taskbar; it is sized to the usable work area so the Next/Back bar is
  always visible on any screen size. (Resolves the roadmap's
  zoomed-vs-fit_window question: fit_window wins.)
- Advanced Query Builder: every "Apply as ..." now creates a visible,
  assignable metric in the sidebar's Saved Queries section — named with the
  typed metric name, or a readable auto-name ("Query: Impressions (sum)")
  when the name is blank. Previously an unnamed apply armed an invisible
  selection and appeared to do nothing. Saved entries also re-arm their
  stored query when clicked again, and the applied metric shows highlighted.
- UI performance with large imports:
  - KPI totals/flags are computed in the background export pass instead of
    on the Tk thread when the review screen opens.
  - The review screen builds per-campaign detail rows lazily on first
    expand instead of pre-building thousands of hidden widgets.
  - The client wizard debounces its search box (one list rebuild 250 ms
    after typing pauses, instead of a full rebuild per keystroke).
  - The Advanced Query Builder caches its data scan per mapper session
    instead of rescanning every level row each time it opens.
- Confirmed the review screen's data flags are computed from the data
  (zero-value metrics and rate-mis-alias heuristics in engine.kpi), not
  hardcoded.
- Query pivot correctness (tables/charts that ship in reports): the pivot
  no longer pools rows across breakdown types. With no breakdown type
  selected it now shows the documented clean campaign-totals table —
  previously every type's rows were pooled and each type's "Other" bucket
  summed into one giant bogus top row. A value present in more than one
  selected type is kept as separate disambiguated rows ("Other (zone)",
  "Other (dow)") instead of being silently summed. The pivot logic moved
  to a pure `build_pivot()` function locked by regression tests.
- Saved Queries render directly below Quick Fill in the mapper sidebar
  instead of after all breakdown sections.
- Excel search fixed for re-exports: the writer no longer edits an
  existing .xlsm in place with openpyxl (which silently dropped the
  ActiveX search box and corrupted the sheet/VBA wiring — the search broke
  on every re-export of the same period). Every export now harvests the
  previous rows, rebuilds the workbook from scratch, and re-injects the
  VBA search engine; merge semantics (same-key replacement, distinct
  periods preserved) are unchanged. A stale macro workbook is never left
  behind holding old data.
- KPI accuracy against vendor dashboards (traced on a real order where
  impressions matched the vendor to the digit):
  - Completion Rate now uses Video Starts as the denominator when the
    export carries a starts metric (vendor VCR definition), falling back
    to impressions otherwise. New "Video Starts"/"Starts" aliases.
  - Cross-campaign Reach/Frequency cannot be deduplicated from campaign
    aggregates (vendor reach was 3.2x lower than the campaign sum). The
    review totals are now labeled "Combined Reach (not deduplicated)" and
    "Avg Campaign Frequency", each with a data flag pointing to the vendor
    dashboard for order-level numbers. Per-campaign values are unchanged.
- 249 automated tests pass.

## July 13, 2026 - MappingModel extraction (mapper roadmap Phase 3)

- Added `app/mapper/mapping_model.py`: `MappingModel` is now the single
  owner of template-mapping state (assignments, image mappings, skip flags,
  per-metric format preferences). Every mapper mutation goes through it;
  the Tk shape panel and the live COM preview are observers that re-render
  from the model after each change, eliminating the four-way state sync
  (widgets / shape dicts / COM working copy / JSON).
- No schema change: the model keeps the exact persisted mapping format,
  including the legacy shape-level single-assignment form. Old mappings
  load unchanged and `engine.pptx_fill` is untouched.
- Assignment semantics preserved verbatim: same-metric re-assign updates in
  place, a full-text assign onto a mapped shape still asks before replacing,
  and format changes propagate to existing assignments without re-assigning.
- 15 new unit tests for the model (schema contracts, assignment semantics,
  observer events, fill-engine integration); 236 automated tests pass.
- The mapper's Windows COM live-preview flows still require the acceptance
  drill in `reviews/MAPPER_RELIABILITY_ROADMAP_2026-07-12.md` before a
  release build.

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
