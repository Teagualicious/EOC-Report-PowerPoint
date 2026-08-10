> **Review context:** This report describes the code review performed before the later folder cleanup. File paths in the report may use the former root-level layout; current paths are documented in `../MODEL_HANDOFF.md`.

# Spectrum Reach Reporting Ingestion Engine
## Code Review, Performance Hardening, and UI Refresh

**Review date:** July 10, 2026  
**Reviewed source:** `IngestionEngine1 38.zip`  
**Reviewed application:** Python/Tkinter desktop ingestion, Excel reporting, and PowerPoint template automation

## Executive summary

The application already had a useful modular structure and a strong baseline test suite. The original build compiled successfully and all **143 original tests passed**, so the review focused on latent correctness problems, large-file efficiency, filesystem safety, Windows reliability, and the user experience rather than repairing an obviously broken project.

The reviewed build now passes **161 tests**, compiles cleanly, launches successfully in a graphical smoke test, and includes a modernized multi-step UI. The largest practical improvements are:

- Large CSV/XLSX imports no longer perform a redundant unified-row conversion before the saved platform mapping performs the same work again.
- Excel header detection is bounded and sheets are streamed instead of materialized twice.
- Re-exporting an existing report replaces matching rows instead of silently doubling metric values.
- Cross-source totals now choose the best source **per campaign**, avoiding dropped campaigns.
- Ratio metrics such as CTR, CPM, CPC, and Frequency are averaged instead of summed in all reviewed report paths.
- CSV files with a UTF-8 BOM retain the correct first header.
- Long-running parsing and Excel generation run off the Tk main thread, so the interface remains responsive.
- User-controlled file and folder names are sanitized, made collision-safe, and protected against path traversal.
- Template bundle import validates archive paths, format/version, file count, expanded size, image references, and duplicate image names.
- The primary workflow, client assignment, and review screens now share a cleaner visual system and step-based navigation.

**Overall assessment:** suitable as a substantially safer and more polished candidate build, with one important qualification: the Windows-only Excel/PowerPoint COM paths still require a final acceptance run on a Windows machine with Microsoft Office installed.

---

## Scope and validation

The review covered the complete Python project, including configuration, parsers, data pipeline, KPI logic, Excel generation, PowerPoint mapping/filling, Tkinter UI, tests, launch scripts, and bundled documentation.

Validation performed:

- Python compilation: **passed**
- Original automated suite: **143/143 passed**
- Expanded reviewed suite: **161/161 passed**
- Graphical launch smoke test under a virtual display: **passed**
- Synthetic large-file benchmark: **passed**
- ZIP/template malicious-path regression tests: **passed**
- Repeat-export idempotency regression test: **passed, including values rather than row count alone**

The test suite gives strong coverage to the pure data/business modules. Windows COM automation and interactive Tk screens are inherently less covered in this Linux review environment.

---

## Performance results

A 60,000-row, five-column synthetic export produced 180,000 normalized metric rows. Results below are the median of three isolated process runs. Peak memory is maximum resident set size.

| Main workflow | Original | Reviewed | Improvement |
|---|---:|---:|---:|
| CSV elapsed time | 2.03 s | 1.86 s | 8.4% faster |
| CSV peak memory | 406,544 KB | 369,352 KB | 9.1% lower |
| Excel elapsed time | 6.96 s | 4.85 s | 30.3% faster |
| Excel peak memory | 421,168 KB | 371,048 KB | 11.9% lower |

Why Excel improved most:

1. Header detection now scans only the first 250 rows rather than requiring the entire sheet in memory.
2. Worksheet rows are streamed after header detection.
3. The main workflow uses a table-only parser mode and lets the platform mapping build unified rows once, rather than constructing and then discarding a duplicate set.
4. Workbooks are closed deterministically in a `finally` block.

There is still a structural memory floor because the platform-mapping design retains source table rows while also creating normalized rows. A future architecture could stream directly through a known platform configuration and retain only a small preview of the raw table.

---

## Correctness and bug fixes

### 1. Repeat exports could double values

**Severity: High**

The previous merge behavior added an existing row and a newly imported row when all keys matched. Running the same monthly export twice could therefore double impressions, clicks, cost, and other metrics while leaving the row count unchanged.

**Fix:** duplicates within one new import batch are still aggregated correctly, but a new exact-key row now replaces the prior workbook row. The regression test compares every exported row value across two runs.

### 2. Best-source selection could drop campaigns

**Severity: High**

The query resolver previously selected one globally largest source for an entire metric. If one campaign's best data came from a ZIP breakdown and another campaign's best data came from a campaign summary, one campaign could disappear from the total.

**Fix:** additive totals now choose the best source per campaign and then sum campaigns. Ratio metrics prefer a campaign summary when available and otherwise choose one representative source per campaign before averaging.

### 3. Acronym metrics were changed during export

**Severity: High**

Title casing converted `CTR` to `Ctr` and `CPM` to `Cpm`. That could break exact dictionary lookups, aggregation rules, searches, and PowerPoint mappings.

**Fix:** universal metric names are preserved exactly. Client, campaign, and source display text can still be normalized independently.

### 4. Ratio metrics could be summed

**Severity: High**

The structured mapper preview could sum values such as CTR or Frequency even when the inserted PowerPoint value was averaged. Duplicate ratio rows in Excel were also treated like additive metrics.

**Fix:** all reviewed paths consult the metric dictionary. Additive metrics sum; ratio metrics average. Preview and inserted values now agree.

### 5. UTF-8 BOM could corrupt the first CSV header

**Severity: Medium**

Trying ordinary UTF-8 before UTF-8-with-signature could leave a BOM attached to `Campaign`, preventing column-role recognition.

**Fix:** `utf-8-sig` is attempted first and the file is streamed after bounded encoding/delimiter detection.

### 6. Legacy `.xls` was offered but unsupported

**Severity: Medium**

The file picker advertised `.xls`, but the parser uses `openpyxl`, which does not read legacy binary `.xls` workbooks.

**Fix:** the UI now advertises XLSX/XLSM only, and parser errors clearly instruct users to save legacy files as XLSX first.

### 7. Filtered client selections disappeared

**Severity: High, workflow usability**

The client wizard rebuilt its checkbox variables whenever the user typed in the search box. Previously selected campaigns could silently become unselected or be omitted from assignment.

**Fix:** selection state is persistent across filtering and rebuilding. The client wizard also prevents a duplicate client name from overwriting an earlier assignment.

### 8. Metric-only HTML reports could be erased by configuration

**Severity: High**

Platform configuration could rebuild from detected tables and replace parser output with empty data when no table matched. Metric-only HTML dashboards were particularly vulnerable.

**Fix:** flat HTML metrics are exposed as a one-row `HTML Report` source with the parsed campaign name as a fallback. If no configured sheet matches at all, parser output is preserved rather than silently erased.

### 9. Date-range regex was overly permissive

**Severity: Low/Medium**

A character class was used where alternation was intended, allowing invalid single letters to be interpreted as a date separator.

**Fix:** separators are explicit alternatives: hyphen, en dash, em dash, or the word `to`.

### 10. Client output names could collide

**Severity: Medium**

Names such as `A/B`, `A:B`, and `a_b` can sanitize to the same Windows folder. This could merge or overwrite two clients' exports.

**Fix:** output components are made unique using case-insensitive collision tracking and numeric suffixes.

---

## Security and resilience hardening

### User-controlled names

Platform names, fingerprint names, template mappings, template renames, image names, and client output folders now use centralized Windows-safe naming helpers. Invalid characters, path separators, control characters, trailing dots/spaces, `..`, and reserved device names such as `CON` are neutralized.

### Atomic persistence

Settings, platform configurations, HTML fingerprints, and PowerPoint mappings now write to a temporary file, flush to disk, and replace the destination atomically. A crash during a write is far less likely to leave truncated JSON.

### Template bundle import

Imported ZIP bundles now enforce:

- recognized bundle format and version;
- safe template basename ending in `.pptx`;
- no absolute paths or `..` archive paths;
- maximum archive file count;
- maximum total expanded size;
- supported image types;
- no case-insensitive duplicate image basenames;
- all tokenized image references must exist;
- atomic replacement of the template file;
- overwrite refusal unless explicitly requested.

This protects against accidental corruption and common ZIP path traversal problems.

---

## UI refresh

The main workflow was redesigned around a consistent visual system rather than isolated widget changes.

### Main import screen

- Branded navy application header
- Three-step progress indicator: Files → Clients → Review
- Clear report-building heading and explanatory copy
- Card-based source-file area
- Scrollable file list with empty state
- File count badge
- Consolidated reporting-period card
- Prominent primary action
- Status line and indeterminate progress indicator
- Responsive minimum window size
- Native Windows DPI-awareness initialization

### Client assignment

- Shared branded header and step indicator
- Card-based client/search controls
- Persistent selections while filtering
- Modern primary/secondary buttons
- Duplicate-client protection

### Review

- Shared branded header and current-client context
- Refreshed KPI cards using the theme's card/border tokens
- Consistent report, next, finish, and back actions
- Existing expandable campaign detail and data-flag interactions retained

### Interaction improvements

- Parsing and workbook export execute on background threads.
- Tk callbacks and UI updates remain on the main thread.
- Date fields validate format and order before processing.
- Platform changes in Settings refresh all existing file-row dropdowns.
- File lists scroll correctly with high-resolution mouse wheels.
- Template rename validates names and prevents collisions.

---

## Maintainability findings

The overall module boundaries are reasonable, but several functions remain too large and should be decomposed in a future refactor:

| Function | Approximate size | Concern |
|---|---:|---|
| `mapper/query_builder.py::show_query_builder` | 402 lines | UI state, query construction, pivot refresh, and event wiring are combined |
| `ui/review_view.py::_show_review_in_main` | 211 lines | layout, KPI formatting, flags, details, and navigation are combined |
| `ui/review_view.py::_generate_report` | 170 lines | selector UI and generation orchestration are combined |
| `mapper/slide_view.py::_assign_to_shape` | 165 lines | assignment validation and multiple mapping types are combined |
| `engine/query_resolver.py::resolve_query` | 155 lines | string tokens, data extraction, source choice, filters, and aggregation are combined |
| `engine/metrics_catalog.py::get_available_metrics` | 141 lines | collection, aggregation, flat catalog, and structured preview are combined |

Recommended direction:

1. Create one canonical function that converts `client_data` to a normalized DataFrame. KPI, metric catalog, and query resolver currently repeat similar row-collection logic.
2. Separate query parsing, source selection, filtering, and aggregation into pure functions.
3. Break Tk screens into component builders and state/controller methods.
4. Introduce typed dataclasses or `TypedDict` definitions for parsed data, mappings, and export results.
5. Narrow broad `except Exception` blocks inside core logic while retaining broad boundaries at UI/task entry points for user-friendly reporting.

---

## Remaining risks and recommended next steps

### 1. Windows Office automation acceptance test

**Priority: Required before release**

This environment cannot execute Microsoft Excel or PowerPoint COM automation. The following modules remain the highest integration risk despite compiling successfully:

- `engine/excel_vba.py`
- `engine/pptx_fill.py`
- `engine/pptx_live.py`
- `engine/pptx_mapper.py`
- `engine/pptx_thumbs.py`

Run the supplied manual checklist on Windows 10/11 with the supported Office build. At minimum test:

- VBA injection and `.xlsx` → `.xlsm` upgrade
- re-export into an existing `.xlsm`
- locked workbook behavior
- PowerPoint live preview startup/shutdown
- text, image, chart, and table shape fills
- template thumbnail generation
- PowerPoint process cleanup after errors

### 2. Automated UI tests

**Priority: High**

Pure logic is well tested, but Tk screens and mapper interactions have little automated coverage. Add a small Windows CI smoke suite or a GUI harness covering:

- add/remove/clear files;
- platform dropdown refresh;
- client filtering with persistent selections;
- date validation;
- background task completion/error callbacks;
- review navigation;
- template selector and mapper opening.

### 3. Cancellation and progress granularity

**Priority: Medium**

Background work prevents freezing, but users cannot cancel a large import or Office export. Add a cancellation event checked between files/sheets and report determinate progress where row counts are available.

### 4. Canonical data model

**Priority: Medium**

KPI, catalog, query, and Excel logic still have parallel aggregation implementations. Consolidating them would reduce the chance that a future metric type behaves differently between review, Excel, and PowerPoint.

### 5. Continuous integration and packaging verification

**Priority: Medium**

Add a clean-environment CI job that installs from the declared requirements, runs compile/tests, builds the executable/portable distribution, and verifies that required assets/templates are included.

---

## Files changed at a glance

Core additions and major edits include:

- `config/naming.py` — safe and unique filesystem components
- `config/settings.py` — atomic persistence and safe platform paths
- `config/themes.py` — modern light/dark design tokens
- `parsers/csv_parser.py` — BOM-safe streaming and table-only mode
- `parsers/excel_parser.py` — bounded header scan, streaming, table-only mode, deterministic close
- `parsers/html_parser.py` — faster fingerprints, safer persistence, encoding fallback, corrected date matching
- `engine/data_pipeline.py` — non-duplicative parse path, HTML fallback, preserve unmatched output
- `engine/excel_utils.py` — correct replacement/idempotency and ratio aggregation
- `engine/query_resolver.py` — per-campaign source resolution
- `engine/metrics_catalog.py` — consistent ratio preview behavior
- `engine/pptx_mapper.py` — safe mapping paths and atomic writes
- `engine/template_bundle.py` — hardened portable bundle import/export
- `ui/utils.py` — shared modern components and background-task delivery
- `ui/main_window.py` — responsive modern workflow and background processing
- `ui/client_wizard.py` — modernized assignment flow and persistent filtering state
- `ui/review_view.py` — consistent review header/cards/actions
- `ui/settings_window.py` — safe template rename
- `main.py` — Windows DPI awareness
- expanded regression suite, including new naming/path-safety tests

## Release recommendation

Use this reviewed build as the new development baseline. Complete one Windows/Office acceptance pass before distributing it broadly, because those integrations cannot be proven by the Linux test environment alone.
