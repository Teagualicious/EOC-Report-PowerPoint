# AI Context — Spectrum Reach Reporting Ingestion Engine

**Purpose:** Fast, reliable orientation for an AI assistant making changes to this repository.  
**Authoritative state:** July 13, 2026  
**Read next:** `documentation/MODEL_HANDOFF.md` for the detailed system handoff. For repo session workflow (read `STATUS.md` first, tests as the completion gate, synthetic fixtures only), follow root `CLAUDE.md` and `STATUS.md`.

## 1. Project intent

This is a local-first Windows desktop application that converts advertising-platform exports into normalized, searchable Excel workbooks and mapped PowerPoint reports. It is designed for nontechnical users who need a guided monthly workflow rather than a command-line data tool.

The product priorities, in order, are:

1. **Correct reporting totals.** Never trade accuracy for a cosmetic or performance improvement.
2. **Safe repeatability.** Re-importing the same reporting period must not silently double data.
3. **Simple user navigation.** Inputs and outputs are visible at the project root; internal state remains separated.
4. **Responsive desktop behavior.** Parsing and report generation must not freeze Tkinter.
5. **Graceful Windows integration.** Excel/PowerPoint automation should enhance the result, but failures must produce clear warnings rather than data loss.
6. **Portable handoff.** Another developer or model should be able to understand, test, and extend the project without reverse-engineering prior conversations.

## 2. Current repository boundaries

```text
Jughead-Data-Engine/
|-- AI_CONTEXT.md                  AI-specific orientation (this file)
|-- CLAUDE.md                      Repo session rules
|-- STATUS.md                      Current project state; read first each session
|-- README.md                      User/developer starting point
|-- requirements.txt               Dev/CI install (pulls in app/requirements.txt plus pytest)
|-- Start Ingestion Engine.bat     Normal Windows launcher
|-- app/                           Application code and bundled resources
|-- input/                         User-facing source-file staging
|-- output/                        Default generated-report destination
|-- workspace/                     Writable settings, mappings, templates, logs
|-- documentation/                 Current guides and historical archive
|-- tests/                         Pytest suite (synthetic fixtures only)
`-- developer/                     Build configuration and tools
```

Respect these boundaries:

- `app/` is source code and read-only packaged resources.
- `input/` and `output/` are intentionally at the root for convenient File Explorer access.
- `workspace/` is application-managed state. Do not put source modules or user reports there.
- `documentation/archive/` is historical context, not current truth.
- `tests/` uses synthetic fixtures only (see `tests/fixtures/`); never real client data.
- `developer/` is not required for a normal end-user run.

All runtime paths must come from `app/config/paths.py`. Do not add ad-hoc `__file__`, current-working-directory, or hard-coded absolute paths.

## 3. Commands to run before and after changes

From the project root:

```bash
python -m pip install -r requirements.txt
python app/main.py
python -m pytest tests -q
python -m compileall -q app tests
```

Current verified baseline: **262 tests pass**. GitHub Actions (`.github/workflows/ci.yml`) runs the suite on every push and pull request (Python 3.12, ubuntu). The Windows launcher still installs from `app/requirements.txt`; the root file includes it and adds pytest.

For Windows releases, also complete the Office acceptance checks in `documentation/TESTING_AND_RELEASE.md`. Linux/macOS test success does not validate Excel VBA injection or PowerPoint COM behavior.

## 4. Architecture map

### Entrypoint and configuration

- `app/main.py` — startup, logging, DPI awareness, exception hooks, directory creation, UI launch.
- `app/config/paths.py` — sole source of path constants and legacy-layout migration.
- `app/config/settings.py` — atomic settings/platform-config persistence.
- `app/config/logging_setup.py` — rotating-file logging to `workspace/logs/`; owns all handlers.
- `app/config/naming.py` — safe, collision-resistant filesystem names.
- `app/config/themes.py` — light/dark design tokens.

### Parsing and normalization

- `app/parsers/csv_parser.py`
- `app/parsers/excel_parser.py`
- `app/parsers/html_parser.py`
- `app/parsers/dictionary.py` — metric aliases, roles, levels, and aggregation rules.
- `app/resources/metric_dictionary.json` — business dictionary; modify cautiously and test aliases/aggregation.
- `app/engine/data_pipeline.py` — platform setup scan, configured parsing, campaign filtering.

### Report calculations and exports

- `app/engine/kpi.py` — review totals.
- `app/engine/metrics_catalog.py` — values exposed to PowerPoint mapping.
- `app/engine/excel_writer.py` — workbook generation and merge/re-export behavior.
- `app/engine/excel_utils.py` — UI-free normalization/collection/pivot helpers for the Excel writer.
- `app/engine/excel_search_dashboard.py` — Search sheet and hidden index/config sheets.
- `app/engine/excel_vba.py` + `app/engine/vba_src/modSearch.bas` — Windows Excel enhancement.
- `app/engine/query_resolver.py` — advanced metric selection/filtering.
- `app/engine/pptx_mapper.py` — template scanning and mapping storage; re-exports the fill/format/catalog companions.
- `app/engine/pptx_fill.py` — static `python-pptx` output; `fill_template_report()` returns `(path, report)`.
- `app/engine/fill_report.py` — fill telemetry: `FillReport` (missing metrics, unmatched `replace_text` placeholders, missing images) appended to `workspace/logs/fill_history.jsonl`.
- `app/engine/pptx_live.py` — Windows COM preview and complex PowerPoint updates; tracks COM health and self-disables after 3 consecutive failures, falling back to the `python-pptx` engine.
- `app/engine/pptx_thumbs.py` — cached slide-1 template previews (COM PNG on Windows, text summary fallback).
- `app/engine/pptx_formats.py` — shared formatting behavior.
- `app/engine/template_bundle.py` — safe template import/export ZIP handling.
- `app/engine/errors.py` — structured application errors with user-safe messages.

### UI

- `app/ui/main_window.py` — primary workflow and application state.
- `app/ui/client_wizard.py` — campaign-to-client assignment.
- `app/ui/review_view.py` — review and generation flow.
- `app/ui/settings_window.py` — platforms, appearance, templates, logs.
- `app/ui/platform_setup.py` — mapping platform export columns.
- `app/ui/utils.py` — approved shared widgets/styles/background-task helper.
- `app/mapper/mapping_model.py` — `MappingModel`, the single owner of template-mapping state; all mapper mutations go through it and observers re-render from it.
- `app/mapper/` — PowerPoint mapper window, sidebar, slide view, format/query dialogs (render and gather input only — mapping mutations belong in the model).

### AI-facing interfaces (headless, no Tk)

- `app/engine/workflow.py` — the shared end-to-end workflow service (parse → campaigns → KPIs → export → fill → query); the UI and both AI interfaces must stay thin over it.
- `app/cli.py` — JSON-in/JSON-out terminal interface for agents and automation.
- `app/mcp_server.py` — local MCP server exposing the workflow as Claude tools (`mcp` package required only on the hosting machine; `INGESTION_MCP_READ_ONLY=1` disables the writing tools). See `documentation/AI_INTEGRATION.md`.

## 5. Non-negotiable business invariants

### Metric aggregation

Use `parsers.dictionary.get_metric_aggregation()` as the source of truth.

- Additive counts and currency generally use `sum`.
- Rates and ratios use `avg`.
- Never sum CTR, CPM, CPC, VCR, Frequency, or equivalent ratio metrics unless the business dictionary explicitly changes.

### Acronym preservation

Universal names such as `CTR`, `CPM`, `CPC`, and `VCR` must remain uppercase. Avoid generic title-casing that turns them into `Ctr` or `Cpm`.

### Best-source selection

Campaign totals can appear in multiple source/breakdown tables. For additive metrics, select the most complete source **per campaign and metric**, then sum the campaign winners. Do not select one global source for an entire client.

### Repeat exports

An existing output workbook is mergeable state:

- Same complete identifying key: replace/update the old row.
- Different date/source/campaign/level/metric: preserve separately.
- Duplicate rows inside one incoming batch: collapse using the metric aggregation rule.

Any change to Excel merge keys needs explicit tests for idempotency and distinct-period preservation.

### Parsed-data contract

Parsers and configured mappings must preserve the dictionary shape documented in `documentation/MODEL_HANDOFF.md`, especially:

- `source_file`, `source_platform`
- `campaign_name`, `client_name`
- `start_date`, `end_date`
- `campaign_metrics`
- `level_data`
- `mapped_metrics`, `metrics`, `detected_tables`

Do not casually rename these keys; they are shared across UI, KPI, Excel, and PowerPoint code.

## 6. Threading and UI rules

Tkinter is single-threaded.

- Long parsing, workbook generation, and report-generation work must use `ui.utils.run_in_background()` or an equivalent pattern already used by the app.
- Worker threads must not read or mutate Tk widgets.
- UI updates belong in callbacks scheduled on the Tk event loop.
- Always restore button/progress state on both success and error paths.
- Do not introduce blocking network calls; the application is intentionally local-first.

UI changes should reuse theme tokens from `config/themes.py` and helpers in `ui/utils.py`. Avoid isolated hard-coded colors, inconsistent padding, or new widget styles that only work in one theme.

Window geometry is DPI-aware. Use `ui.utils.fit_window()` for every new
top-level window, avoid fixed-height headers, and place expandable canvases in
their own container rather than mixing `pack(side="left")` with later bottom
action bars. The latter pattern caused controls to overlap on Windows laptops.

## 7. Files that require extra caution

Changes here can have broad or platform-specific effects:

- `app/config/paths.py` — affects migration, portable builds, source runs, and all user data locations.
- `app/resources/metric_dictionary.json` — changes matching and aggregation across every parser.
- `app/engine/excel_writer.py` — affects report correctness and repeat-export behavior.
- `app/engine/excel_search_dashboard.py` and `app/engine/vba_src/modSearch.bas` — must remain synchronized.
- `app/engine/pptx_live.py` — Windows COM lifecycle and Office-version sensitivity.
- `app/engine/query_resolver.py` — advanced mapping semantics.
- `app/ui/main_window.py` — owns shared workflow state; regressions can cross multiple screens.
- `app/config/settings.py` and mapping schemas — must remain backward-compatible where practical.

“Extra caution” does not mean never modify these files. It means read their tests and dependent modules first, preserve compatibility, and add regression coverage.

## 8. Persistence and compatibility philosophy

- JSON writes for settings and mappings must remain atomic.
- Existing user files win during migration; never overwrite them silently.
- Legacy migration is non-destructive and idempotent.
- Preserve custom output folders. Only known former defaults should be translated automatically.
- Prefer additive schema evolution: tolerate missing old fields and supply defaults.
- Template mappings can contain legacy relative image paths; keep resolution compatibility unless there is a documented migration.
- Sanitized names can collide. Use `config.naming` helpers rather than hand-written character replacement.

## 9. Security expectations

Although this is a desktop application, treat imported files and template bundles as untrusted.

- Keep ZIP traversal, archive-size, required-member, and filename validations in `template_bundle.py`.
- Do not evaluate formulas, execute macros from imported files, or run shell commands derived from user content.
- Do not log sensitive report contents unnecessarily.
- Avoid exposing raw exception traces to normal users; write detail to rotating logs and show actionable messages.
- Maintain file-extension validation and clear guidance for unsupported legacy `.xls` files.

## 10. Performance assumptions

The reviewed build was optimized for large tabular exports. Preserve these patterns:

- Avoid repeated full-DataFrame copies.
- Select/read only required columns when possible.
- Normalize headers once rather than repeatedly per row.
- Prefer vectorized operations or bounded iteration over nested scans.
- Do not load an entire workbook multiple times in one workflow without a measured reason.
- Keep UI rendering proportional to visible items; use scrollable containers for long lists.

Before claiming a speed improvement, benchmark representative CSV and XLSX files and confirm output equivalence. A faster wrong total is a regression.

## 11. Coding conventions

- Target Python 3.11+ (CI runs 3.12).
- Use descriptive names and small focused functions.
- Add type hints to new public functions where practical.
- Catch specific exceptions; do not hide unexpected failures with broad silent `except` blocks.
- Use `logging`, not `print`, for runtime diagnostics.
- Keep user-facing language concise and actionable.
- Centralize reusable formatting, naming, and path behavior.
- Write regression tests for every fixed bug.
- Avoid new dependencies unless the benefit is substantial and Windows installation/packaging is considered.

## 12. Documentation maintenance rule

A code change is incomplete when it changes behavior described in documentation but leaves the docs stale.

Update as applicable:

- `AI_CONTEXT.md` — design constraints, pitfalls, preferred patterns.
- `documentation/MODEL_HANDOFF.md` — authoritative implementation handoff.
- `documentation/CURRENT_ARCHITECTURE.md` — repository/module structure.
- `documentation/API_REFERENCE.md` — public contracts and schemas.
- `documentation/USER_GUIDE.md` — user-visible workflow.
- `documentation/TECHNICAL_GUIDE.md` — technical behavior.
- `documentation/TESTING_AND_RELEASE.md` — commands/acceptance tests.
- `documentation/CHANGELOG.md` — meaningful released changes.
- `documentation/PROJECT_MANIFEST.json` — commands, status, critical docs, hashes.

The generated PDF guides were removed; the Markdown files are canonical.

## 13. Preferred change workflow for an AI assistant

1. Read this file and `documentation/MODEL_HANDOFF.md`.
2. Inspect the directly affected module, its callers, and its tests.
3. State assumptions explicitly when business behavior is not encoded.
4. Make the smallest coherent change that solves the problem.
5. Add or update regression tests.
6. Run the focused tests, then the full suite and compile check.
7. Update relevant documentation and manifest metadata.
8. For Office features, clearly distinguish automated validation from outstanding Windows acceptance testing.
9. Deliver a clean archive without caches, temporary reports, credentials, or local machine paths.

## 14. Common failure modes to avoid

- Using the current working directory as the project root.
- Moving `input/` or `output/` back under `workspace/`.
- Updating code paths but forgetting the launcher, tests, docs, or migration.
- Summing rate metrics.
- Deduplicating only by metric name and losing distinct dates/campaigns/levels.
- Choosing one global “best” source and dropping campaigns from other sources.
- Touching Tk widgets from a worker thread.
- Assuming passing Linux tests validates Office COM behavior.
- Breaking old JSON mappings by requiring newly added fields.
- Replacing PowerPoint text in a way that destroys surrounding runs/formatting.
- Trusting ZIP member paths or imported filenames.
- Adding root-level clutter instead of placing code, state, docs, or tools in their assigned directory.

## 15. Known remaining risk and technical debt

- Excel VBA injection and PowerPoint COM need final Windows/Office acceptance testing on the target environment.
- COM automation remains inherently sensitive to Office version, security settings, active dialogs, and orphaned processes.
- Tkinter UI modules are still relatively large and could later be decomposed into smaller state/controller/view components, but a broad rewrite is higher risk than incremental extraction.
- Parser support depends on platform export formats; new vendor layouts should be added through platform configuration and regression fixtures rather than one-off UI logic.
- Additional benchmark fixtures using real anonymized exports would improve confidence beyond synthetic performance tests.

## 16. High-value future improvements

Only pursue these with tests and a clear user benefit:

1. Add a first-run diagnostics screen for Python, dependencies, write permissions, Excel, PowerPoint, and Trust Center configuration.
2. Add anonymized golden-file integration fixtures for complete CSV/XLSX → Excel/PPTX workflows.
3. Add schema/version fields and migrations to settings and mapping JSON.
4. Add cancellation support and richer progress reporting for very large imports.
5. Add structured validation summaries before generation (missing campaign assignments, date gaps, unmapped metrics).
6. Gradually extract large Tkinter modules without altering the visible workflow.
7. Add Windows CI or a repeatable Office acceptance harness where licensing/environment permits.

## 17. Definition of done

A change is ready to hand off when:

- The intended user behavior works.
- Reporting invariants remain correct.
- Relevant regression tests exist.
- Full tests and compile checks pass.
- UI remains responsive for long operations.
- Existing settings/mappings/data remain usable or migrate safely.
- Current documentation reflects the change.
- Office-specific limitations are tested on Windows or explicitly listed as outstanding.
- The packaged archive is clean and launches from `Start Ingestion Engine.bat`.
