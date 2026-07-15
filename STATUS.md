# STATUS

> The single source of truth for project state. Every session reads this first and updates it last.

## Current phase

Phase 8 — Demo readiness (complete): portable release download + client-data hygiene scrub

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
  - 15 new model unit tests; 236 tests pass; `python -m compileall -q app tests` clean. All 221 pre-refactor tests pass unchanged (the golden-suite acceptance gate). Released as v1.24.0.
- Phase 3 (2026-07-13) — Windows debugging batch (from on-site photos):
  - Mouse wheel can no longer change a closed combobox's value anywhere (class-binding override in `ui.utils.enable_mousewheel`; wheel scrolls the surrounding pane instead). This was silently corrupting Platform Setup role selections.
  - Platform Setup uses one grid per sheet so Role/Sample columns align across rows.
  - Client wizard no longer zooms over the taskbar; `fit_window(1280, 940)` clamps to the work area so Next/Back are always visible (resolves the roadmap's zoomed-vs-fit_window question in favor of fit_window).
  - Advanced Query Builder "Apply as ..." always creates a visible sidebar metric (typed name or auto-name), re-armable from Saved Queries; armed selection renders highlighted.
  - Verified review data flags are computed (engine.kpi zero-value + rate-mis-alias heuristics), not hardcoded.
  - Performance: KPI compute moved into the background export pass; review detail rows build lazily on expand; client-wizard search debounced (250 ms); query-builder data scan cached per mapper session.
  - Round 2 (Windows re-test feedback): Saved Queries moved directly below Quick Fill; **query pivot correctness fix** — the pivot no longer pools rows across breakdown types (each type re-slices the same delivery, so each type's "Other" bucket summed into a giant bogus top row in shipped tables). Empty breakdown selection now yields the documented campaign-totals table; colliding values across selected types render as disambiguated rows ("Other (zone)"). Pivot extracted to pure `build_pivot()` with regression tests.
  - Round 3 (KPI-vs-vendor audit): impressions matched the vendor exactly; three derived metrics explained and addressed. Completion Rate now divides by Video Starts when present (vendor VCR; the impressions-only denominator read several points low on a real order). Reach/Frequency totals cannot be deduplicated from campaign aggregates (a real order's campaign sum ran roughly 3x the vendor's deduplicated total) — relabeled "Combined Reach (not deduplicated)" / "Avg Campaign Frequency" with explanatory data flags; per-campaign values unchanged. Completions ±1 vs vendor traced to best-source picking a breakdown that sums 1 higher than the vendor's own summary (their rounding) — accepted.
  - Round 4 (Excel search regression): root cause was the re-export path editing the existing .xlsm in place with openpyxl, which cannot round-trip macro workbooks (ActiveX search box dropped, sheet/VBA wiring corrupted) — first export worked, every re-export of the same period broke search. `write_to_excel` now harvests existing rows read-only, rebuilds the workbook from scratch, and re-injects the VBA; a stale .xlsm is never left holding old data. Merge semantics unchanged, locked by a lifecycle regression test.
  - 249 tests pass (13 new this phase). Released as v1.25.0.
- Phase 4 (2026-07-13) — AI-native interfaces (Claude Business readiness):
  - `app/engine/workflow.py` — the full workflow (parse → campaigns → KPIs → export → fill → query) as one headless service; UI, CLI, and MCP server all drive it (thin-shell rule enforced by test).
  - `app/cli.py` — JSON terminal interface for agents/automation (8 subcommands, `{"ok":…}` envelope, exit codes, no prompts).
  - `app/mcp_server.py` — local MCP server for Claude Desktop/Code: 8 tools over stdio, data stays on-machine, `INGESTION_MCP_READ_ONLY=1` for analysis-only mode. `mcp` package deliberately NOT in app runtime requirements.
  - `documentation/AI_INTEGRATION.md` — setup, tool table, example prompts, governance notes.
  - 260 tests pass (11 new). MCP server tool registration smoke-verified against the real `mcp` package.
  - Round 2 (search pivot contract): the Excel search now behaves as a strict typed pivot table — columns render in the exact typed order (dims and metrics interleaved; implied dims first, default metrics appended last). KPI summary cards kept but gated (user request): they render only for explicitly typed metrics and only when the table has >1 row — never for the auto-filled default set that produced unrequested boxes/roll-ups. modSearch.bas WriteResults/ParseTerms rewritten around a unified column sequence (mColSeq); grammar docs and Search-sheet hint updated; VBA/dashboard contract locked by static tests. **Needs the Windows search drill re-run** (typed-order searches, e.g. "Impressions, Campaign, Client" and "Client, Campaign, Zip Code"; KPI cards only in the first case).
  - 262 tests pass (13 new this phase).
- Phase 5 (2026-07-14) — Mapper roadmap Phase 4, stable shape identity:
  - Fills resolve each mapped shape by persistent PowerPoint id (`shape_uid`, python-pptx `shape.shape_id` / COM `Shape.Id`) with unique-name fallback; positional index only for legacy entries without stored identity. Kills the SHAPE INDEX DRIFT wrong-shape bug class when templates are edited.
  - Scans (`pptx_mapper`, both scanners) emit `shape_uid` alongside the positional `shape_id` (mapping-JSON keys unchanged). `MappingModel.set_scan_identity()` + lazy stamping put `shape_uid`/`shape_name` on entries as the user touches them; loading/saving old mappings changes nothing (schema byte-identical without a scan attached).
  - Shared pure resolver `app/engine/shape_identity.py` used by `pptx_fill` and new `PPTXLivePreview._resolve_shape` (fast path = one COM Id read; drift retargets by id or skips — no more "warn but write anyway"). Deleted mapped shapes surface via `FillReport.missing_shapes` in the With-Gaps dialog.
  - Tests: `test_scan_and_fill_agree_on_shape_identity` deliberately replaced by reorder/insert/delete drift e2e tests; resolver units (`test_shape_identity.py`); COM stub tests; model stamping tests. 279 pass (17 new); all pre-Phase-4 tests pass unchanged (backward-compat gate).
  - Prompted by the question "would pptxgenjs fix the mapper?" — answered in the plan: pptxgenjs is generation-only (cannot open/edit existing templates) and addresses none of the real issue classes; this refactor is the actual fix.
- Phase 6 (2026-07-14) — Windows debugging batch 3 (from on-site photos):
  - **Output folder is dynamic again.** Saving Settings pinned the absolute default path; after a project-folder rename, exports recreated the old folder (`…-1.27.0\output`) and wrote there. Settings now never persist the default (unset key = follow the project), and `_normalize_settings` falls back to the current `output/` when a stored default-shaped path (basename `output`) no longer exists. Existing custom folders — including ones literally named `output` that still exist — are preserved.
  - **Export stage (client selection → review) sped up + live progress.** One shared hidden Excel process per batch (`engine.excel_vba.ExcelSession`; `inject_search_vba(app=)` / `write_to_excel(excel_app=)`) instead of a full Excel launch per client; Unified Data writer switched from `iterrows` to `itertuples` (row-Series construction was the slow loop). `run_in_background` gained an ordered `on_progress` channel; the status line now narrates each stage per client.
  - **Reach/Frequency hardcoded out of review KPIs** (totals AND per-campaign detail) until real household-level dedup data/formulas exist — supersedes the 2026-07-13 relabel-and-flag rule. Values remain in exported workbook rows and the Excel search.
  - **Consistent KPI number formatting** via `engine.kpi.format_kpi_value` used by cards, campaign summaries, and detail rows: counts render whole with commas (float noise tolerated), rates get `%`, money gets `$`.
  - 291 tests pass (12 net new this phase).
- Phase 7 (2026-07-14) — Windows debugging batch 4 (from on-site stress testing):
  - **Multi-line text fill loss fixed** (the "assigned text doesn't show up in the export" bug, mostly via Auto-Fill Report → built-in engine). Root causes proven at XML level: (a) multi-line replace targets carry `\n` (UI paragraph joins) / `\v` (soft breaks) that run text never contains → per-paragraph match silently no-oped while the report's frame-text check said "filled"; (b) multi-line values written into runs store a literal `\n` in `<a:t>` which PowerPoint renders as whitespace (looked right in COM live preview, flattened in export). Fixes in `engine/pptx_fill.py`: line-by-line target matching (first line gets value, rest cleared, single-line behavior byte-identical), `_explode_newlines` converts written `\n` into real `<a:br/>` (formatting copied), `_replace_in_text_frame` returns the outcome and the report records from it — failed replaces are always unmatched-placeholder now. Full-assign to an empty box writes instead of no-op.
  - **Repeated-run stability**: `PPTXLivePreview.cleanup()` unbalanced the UI thread's COM apartment (CoUninitialize on every call; cleanup runs twice per session via explicit close + `__del__`) → Office automation failed after ~2–3 report cycles. Now `_com_initialized` guard = exactly one CoUninitialize per CoInitialize, idempotent cleanup. `_scan_with_com` closes the presentation and balances COM on failure (try/finally). Thumbnail exports serialized (`pptx_thumbs._EXPORT_LOCK`) so fast template-list clicking can't pile up PowerPoint processes (part of roadmap Phase 5's known issue).
  - 297 tests pass (6 new: multi-line target across paragraphs, soft-break target, failed-replace reporting, real line breaks with formatting, empty-box write, idempotent cleanup).
- Phase 8 (2026-07-15) — Demo readiness:
  - **Portable download**: the Release workflow attaches `IngestionEngine-<version>-portable-win64.zip` (source + `wheelhouse/` of Windows wheels for all runtime deps, built for Python 3.12; pywin32 fetched explicitly since its win32 marker is false on the Linux runner). Launcher precedence: `app\vendor\` present → skip installs entirely (fixed: the check previously ignored vendor and attempted pip offline); `wheelhouse\` present → offline install with internet fallback; else internet as before. Release notes get a "which download" footer. All four distribution shapes documented in TESTING_AND_RELEASE.md.
  - **Client-data hygiene scrub** (owner-approved, pre-demo): real advertiser names in examples/tests replaced with fictional ones ("Acme Appliance Co", "Acme Motors"); real-but-unattributed vendor-audit figures in STATUS/CHANGELOG/test docstrings/kpi comments genericized; stand-in template's example delivery figures changed. Audit found no client exports, credentials, paths, or personal data anywhere tracked. NOTE: previously published GitHub release notes (v1.29.0) still contain two unattributed figures — editable on GitHub if desired; not possible from a remote session.
  - 297 tests pass. Released as v1.31.0.

## Next up

1. Windows/Office acceptance pass — Excel VBA injection, PowerPoint COM live preview, fill-summary dialogs, `fill_history.jsonl`, the refactored mapper, the Phase-4 drift drill (id parity between python-pptx/COM scans, live cut-paste retargeting, legacy-mapping regression), and now batch 3: rename the project folder → export lands in the renamed folder's `output/`; multi-client export launches ONE Excel process; status line narrates stages; review shows no Reach/Frequency and formats counts/%/$ consistently. Per `documentation/TESTING_AND_RELEASE.md` and the checklist in `documentation/reviews/MAPPER_RELIABILITY_ROADMAP_2026-07-12.md`. Nothing COM-related may be declared verified until this passes.
2. Mapper roadmap Phase 5 — small fixes from the July 11 review (template-preview debounce [thumbnail exports are now serialized — Phase 7 — but the selection debounce is still open], client-wizard drag-select dead code, all-caps case-forcing decision, skip-discards-assignments decision — see Noticed).
3. **Template-first mapper (accepted proposal, not scheduled)** — ingest→classify→map→build rework of the mapper around a JSON IR with named slots. Before starting: read `documentation/proposals/TEMPLATE_FIRST_MAPPER_2026-07-15.md` AND `documentation/reviews/TEMPLATE_FIRST_MAPPER_REVIEW_2026-07-15.md` (critique, required design changes — static shapes copied as verbatim XML, uid-based slot reconciliation — phasing A–D, and the reuse map onto existing modules). Prerequisite from Noticed: extend `resolve_query()` to honor builder-query keys.

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
- 2026-07-13 — Releases became push-driven: the Release workflow now fires on any push to main that changes the root `VERSION` file (tag + GitHub Release from the CHANGELOG top entry; already-released versions are skipped). Chosen because remote sessions cannot push tags or call the GitHub API when the connector is unavailable — bumping VERSION inside the PR makes merging the release action. Manual dispatch kept as fallback.
- 2026-07-13 — **v1.25.0 released** (PRs #8/#9, Windows debugging batch + push-driven releases) — first release published by the new VERSION-file trigger, tag on main merge commit 6665fbe.
- 2026-07-13 — **v1.26.0 released** (PR #10, AI-native interfaces: workflow service + CLI + MCP server). GitHub squash-merged a stale PR head (second occurrence — see CLAUDE.md git notes), so the typed-pivot search fix missed this release and ships as **v1.27.0** immediately after.
- 2026-07-13 — **v1.27.0 released** (PR #11, typed-pivot search contract + gated KPI cards).
- 2026-07-14 — pptxgenjs rejected as a mapper replacement — it only generates new decks (no API to open/edit existing .pptx), while the mapper's job is filling existing client templates; the real fix for wrong-shape assignments is stable shape identity (mapper roadmap Phase 4), implemented instead.
- 2026-07-14 — Mapping entries with stored identity that matches nothing in the deck are **skipped and reported** (`FillReport.missing_shapes`), never resolved positionally — writing into the positional slot's new occupant IS the wrong-shape bug. Legacy entries without stored identity keep positional resolution bit-for-bit. Duplicate shape names never match (no guessing).
- 2026-07-14 — **v1.28.0 released** (PR #12, stable shape identity / mapper roadmap Phase 4) — merge verified complete, release auto-published by the VERSION-file workflow.
- 2026-07-14 — Reach/Frequency **omitted** from review KPIs (supersedes the relabel-and-flag rule): even honestly-labeled non-dedup values invited bad comparisons against vendor dashboards. They return when household-level data/formulas exist; workbook rows keep the raw values.
- 2026-07-14 — The default output folder is never persisted in settings.json (unset key = follow the project folder); only user-chosen custom folders are stored. Stale default-shaped paths from renamed/moved projects fall back on load.
- 2026-07-14 — **v1.29.0** — Windows debugging batch 3; releases on merge via the VERSION-file workflow.
- 2026-07-14 — Fill reporting is outcome-based: `_replace_in_text_frame` returns whether it wrote; "filled" is recorded only when text actually changed. The old frame-text containment check could claim filled for replaces that did nothing (multi-line targets).
- 2026-07-14 — Multi-line replace semantics: first target line receives the value, remaining target lines are cleared; single-line targets keep the exact previous matching (plus a trimmed retry for selections with stray whitespace).
- 2026-07-14 — **v1.30.0** — Windows debugging batch 4 (multi-line fills + COM lifecycle); releases on merge.
- 2026-07-15 — Test/example client names must be OBVIOUSLY fictional ("Acme …" family) — real advertiser names from the inherited fixtures were scrubbed pre-demo; keep it that way in new tests and docs.
- 2026-07-15 — **v1.31.0** — portable release download + data-hygiene scrub; releases on merge.
- 2026-07-15 — **Template-first mapper architecture accepted as the v2 direction** (owner proposal): ingest client decks into a JSON IR with named slots, build new decks instead of editing in place. Documented in `documentation/proposals/` + reviewed in `documentation/reviews/TEMPLATE_FIRST_MAPPER_REVIEW_2026-07-15.md`; key review amendments: static shapes are copied as verbatim XML (never rebuilt from schema), charts are their own phase via chart-part cloning, slots reconcile across re-ingests by shape uid, template store lives under `workspace/` (data hygiene). Current mapper remains the production path until template-first survives a real month-end cycle. Not scheduled; doc-only — no release.

## Noticed (not yet acted on)

<!-- Problems spotted mid-task but out of scope. Harvest these periodically. -->
- Launcher and app docs target Python 3.10+, while repo standard is 3.11+ (CI 3.12). Align when there's a reason to touch the launcher.
- `AI_CONTEXT.md` §13/§17 still describe delivering a "clean archive" — stale process language now that the project is a git repo with CI.
- `tests/test_parsers.py::test_real_input_files` scans root `input/` for optional HTML samples — always empty in the repo/CI, so it's effectively a no-op there.
- Text assignments mapped onto shapes without a text frame are silently unreported by `FillReport` (locked as a characterization test); decide whether to surface it.
- Two `tests/test_ui_helpers.py` DPI tests import `ui.utils` and therefore need a Python built with tkinter (GitHub Actions setup-python and standard Windows installs have it; minimal Linux builds may not).
- Toggling a shape's Skip checkbox replaces its whole mapping and silently discards existing assignments (pre-model behavior, deliberately preserved and locked by `test_set_skip_replaces_shape_mapping`). Decide in mapper Phase 5 whether skip should preserve assignments.
- `PPTXWizard.image_paths` is an unused leftover attribute.
- Query-builder queries carry `campaigns`/`sources`/`values`/`top_n` keys, but `engine.query_resolver.resolve_query()` only honors `metric`/`breakdown`/`filter`/`agg` — a re-resolved builder query (e.g. review auto-fill in a later session) can differ from the pivot total shown at apply time. In-session fills use the cached value and are correct. Needs a resolver extension or query translation.
- Opening the mapper computes `get_available_metrics()` on the Tk thread — noticeable but tolerable on large imports; move behind `run_in_background` if it becomes a complaint.
- The mapper catalog still exposes `Total Reach` / `Avg Frequency` keys carrying the non-deduplicated values (renaming them would break saved template mappings). Decide whether deck-side labels should match the review screen's non-dedup labels, or whether those catalog entries should be dropped in favor of manual vendor numbers.
- If any vendor export contains an order-level (deduplicated) reach/frequency summary row, the KPI engine could prefer it over the campaign sum — needs a sample export to confirm the shape.

## How to run

```
pip install -r requirements.txt
pytest
python app/main.py   # launches the Tkinter app (Windows users: Start Ingestion Engine.bat)
```
