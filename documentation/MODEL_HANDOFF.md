# Model and Developer Handoff Guide

**Authoritative state date:** July 10, 2026  
**Project:** Spectrum Reach Reporting Ingestion Engine  
**Runtime:** Python 3.10+ desktop application using Tkinter  
**Primary platform:** Windows 10/11 with Microsoft Excel and PowerPoint

For AI-assisted work, read root `AI_CONTEXT.md` first, then use this document as the detailed authoritative handoff. A human engineer may begin here directly.

## 1. Current status

The reviewed and reorganized build is healthy:

- 171 automated tests pass with `python -m pytest developer/tests -q`.
- All Python files compile.
- The Tkinter application launches successfully in a graphical smoke test.
- Parsing and Excel generation run in background threads so the UI remains responsive.
- Large-file import performance, duplicate-export behavior, metric aggregation, filename safety, and template-bundle validation were hardened in the July 10, 2026 review.
- The only materially unverified area is the Windows-only Office automation path: Excel VBA injection and PowerPoint COM live preview/fill must be acceptance-tested on Windows with Office installed.

Do not treat documents in `documentation/archive/` as current implementation documentation.

## 2. How to run and test

From the project root:

```bash
python -m pip install -r app/requirements.txt
python app/main.py
python -m pytest developer/tests -q
python -m compileall -q app developer/tests
```

On Windows, normal users launch with `Start Ingestion Engine.bat`.

## 3. Repository layout

```text
IngestionEngine/
|-- Start Ingestion Engine.bat
|-- README.md
|-- app/
|   |-- main.py
|   |-- requirements.txt
|   |-- config/
|   |-- parsers/
|   |-- engine/
|   |-- ui/
|   |-- mapper/
|   `-- resources/
|       |-- metric_dictionary.json
|       `-- assets/
|-- input/
|-- output/
|-- workspace/
|   |-- templates/
|   |-- mappings/
|   |-- logs/
|   `-- settings.json
|-- documentation/
`-- developer/
    |-- tests/
    |-- build/
    `-- tools/
```

`app/` is code and bundled resources. `workspace/` is application-managed state, while root `input/` and `output/` are user-facing working folders. Keep those boundaries intact.

## 4. Entrypoint and path rules

`app/main.py` performs these steps:

1. Adds `app/vendor/` to `sys.path` when portable dependencies exist.
2. Enables Windows DPI awareness.
3. Configures rotating logging.
4. Creates the workspace plus root `input/` and `output/` directories, then copies missing data from older layouts.
5. Installs global Python and Tkinter exception hooks.
6. Creates `ui.main_window.IngestionEngine` and starts the Tk event loop.

All paths must come from `app/config/paths.py`. Important constants:

- `PROJECT_ROOT` - folder containing the launcher.
- `APP_DIR` - `PROJECT_ROOT/app` in source mode.
- `RESOURCE_DIR` - `app/resources` or PyInstaller bundled resources.
- `WORKSPACE_DIR` - `PROJECT_ROOT/workspace`.
- `SETTINGS_PATH`, `MAPPINGS_DIR`, `TEMPLATES_DIR`, `IMAGES_DIR`, and `LOGS_DIR` - writable application-state locations under `workspace/`.
- `INPUT_DIR` and `OUTPUT_DIR` - user-facing folders at `PROJECT_ROOT/input` and `PROJECT_ROOT/output`.
- `APP_ROOT` - compatibility alias for `PROJECT_ROOT`; retained for old mapping paths.

Do not reintroduce module-local `__file__` path calculations except for resources that are intentionally colocated with a module, such as `engine/vba_src/modSearch.bas`.

## 5. Main data flow

1. The user adds files and assigns each file to a configured platform.
2. `engine.data_pipeline.scan_file_structure()` inspects a sample export during platform setup.
3. A parser reads each file:
   - `parsers.csv_parser.CSVParser`
   - `parsers.excel_parser.ExcelParser`
   - `parsers.html_parser.HTMLParser`
4. `engine.data_pipeline.apply_platform_config()` rebuilds campaign and breakdown rows using the saved column-role mapping.
5. `ui.client_wizard.ClientWizard` assigns campaigns to client names.
6. `engine.data_pipeline.filter_data_by_campaigns()` creates client-specific parsed data.
7. `engine.excel_writer.write_to_excel()` writes the Search dashboard and Unified Data, then optionally injects VBA and upgrades to XLSM.
8. `engine.kpi.compute_kpis()` calculates review totals.
9. `engine.metrics_catalog.get_available_metrics()` exposes totals, averages, breakdowns, dates, and client name to the PowerPoint mapper.
10. `engine.pptx_fill.fill_template()` or `engine.pptx_live.PPTXLivePreview` fills a mapped deck.

## 6. Parsed-data contract

A parsed data item is a dictionary shaped approximately as follows:

```python
{
    "source_file": "export.xlsx",
    "source_platform": "Architect",
    "client_name": "Client Name",
    "campaign_name": "Fallback Campaign Name",
    "campaign_type": "",
    "start_date": "2026-06-01",
    "end_date": "2026-06-30",
    "campaign_metrics": {
        "Campaign A|Impressions": {
            "value": 50000,
            "universal_name": "Impressions",
            "campaign_name": "Campaign A"
        }
    },
    "level_data": [
        {
            "metric_level": "device:Roku",
            "metric_name": "Impressions",
            "metric_value": 20000,
            "_campaign": "Campaign A"
        }
    ],
    "mapped_metrics": {},
    "metrics": {},
    "detected_tables": []
}
```

Preserve these keys when extending parsers or the platform setup flow.

## 7. Critical business invariants

### Metric names

Universal metric names and aliases come from `app/resources/metric_dictionary.json`. Preserve meaningful acronyms such as CTR, CPM, CPC, and VCR.

### Aggregation

`parsers.dictionary.get_metric_aggregation()` is the source of truth:

- Additive metrics use `sum`.
- Ratio/rate metrics use `avg`.

Do not sum CTR, CPM, CPC, Frequency, or other ratio metrics.

### Best-source rule

Exports may contain campaign totals and multiple breakdown tables that repeat the same metric. For additive totals:

1. Sum within each campaign, metric, and source type.
2. Select the highest/most complete source for each campaign and metric.
3. Sum those campaign-level winners across campaigns.

The choice is per campaign, not one global source for the entire client.

### Re-export rule

When an output workbook already exists:

- A new row with the same full key replaces the old value.
- Different dates, sources, campaigns, levels, or metrics remain separate.
- Duplicate rows inside the same new import batch are collapsed using the metric's configured aggregation.

This prevents a repeated export of the same reporting period from doubling values.

## 8. Persistence formats

### Settings

`workspace/settings.json` stores theme, output folder, and the configured platform name registry.

### Platform mappings

`workspace/mappings/platform_<safe-key>.json` stores sheet and column roles. Platform names are converted through `config.naming.storage_key()`.

### PowerPoint mappings

`workspace/mappings/pptx_<safe-template-key>.json` stores slide/shape assignments. A shape can contain:

- one legacy `metric` assignment;
- multiple entries in `assignments`;
- `format` and `format_details`;
- `replace_text` for partial replacement;
- an advanced `query`;
- `image_path` and optional `image_path_abs`;
- `skip`.

### Templates

PowerPoint templates live in `workspace/templates/`. Image assets live in `workspace/templates/images/`. Template bundles are validated ZIP files containing `manifest.json`, `mapping.json`, one PPTX, and referenced images.

## 9. UI architecture

- `ui.main_window.IngestionEngine` owns application state and the import workflow.
- `ui.client_wizard.ClientWizard` handles client/campaign assignment.
- `ui.review_view.ReviewMixin` renders review screens and launches report generation.
- `ui.settings_window.SettingsWindow` manages platforms, appearance, templates, exports, and debug logs.
- `ui.utils.fit_window()` compensates for Windows display scaling. New top-level windows should use it rather than raw geometry strings, and expandable canvases should live in their own container with action bars reserved above or below them. See `reviews/UI_REVIEW_2026-07-10.md` for the regressions this prevents.
- `ui.utils.run_in_background()` is the approved pattern for long-running work. Background workers must not touch Tk widgets directly; use the success/error callbacks scheduled on the root event loop.
- `mapper.mapper_window.PPTXWizard` combines `SidebarMixin` and `SlideViewMixin` for template mapping.

Keep Tk operations on the main thread.

## 10. Excel integration

`engine.excel_writer.write_to_excel()` always creates:

- `Search`
- `Unified Data`
- hidden `_SearchIndex`
- hidden `_Config`

On Windows, `engine.excel_vba.inject_search_vba()` uses Excel COM to inject `engine/vba_src/modSearch.bas`, add the ActiveX search control and calendar form, and save as XLSM. If COM, Excel, or the Trust Center setting is unavailable, the function returns the XLSX path plus a user-facing warning rather than losing the report.

Do not modify the VBA search grammar without updating both `modSearch.bas` and the search-dashboard index/config generation.

## 11. PowerPoint integration

Static filling uses `python-pptx` in `engine.pptx_fill`. Live preview and complex chart/table updates use Windows COM in `engine.pptx_live`.

Formatting behavior is centralized in `engine.pptx_formats`. Any new display formatting should flow through `_coerce_number()`, `format_with_details()`, or `_format_value()` so the mapper preview and final deck remain consistent.

Partial replacements must preserve surrounding runs and paragraphs. Image replacements must retain the original shape position and size.

## 12. Safety and reliability controls

- Settings and mapping JSON writes are atomic.
- User-controlled filenames pass through `config.naming` helpers.
- Template bundle imports reject path traversal, excessive file counts, excessive expanded size, duplicate image names, unsupported manifests, and missing image references.
- Parser errors use user-facing messages while preserving detailed logs.
- Logs rotate at 1 MB with three backups.
- The legacy layout migration copies without deleting source files or overwriting workspace files.

Do not weaken these controls for convenience.

## 13. Test boundaries

Automated tests cover parsers, dictionary matching, platform application, KPI logic, query resolution, Excel collection/writing, PowerPoint formatting, template bundles, path naming, and smoke pipelines.

Automated tests do not fully exercise:

- real Excel COM/VBA injection;
- real PowerPoint COM live preview;
- all interactive Tkinter workflows;
- arbitrary customer-specific PowerPoint charts and tables.

Use `documentation/TESTING_AND_RELEASE.md` before release.

## 14. Known limitations and next priorities

1. Complete Windows acceptance testing with current Excel and PowerPoint desktop versions.
2. Test comparison charts with multiple series and differing category orientations.
3. Test template image mappings created before and after the workspace restructure.
4. Consider adding a version constant and surfacing it in the Settings/Logs tab.
5. Consider moving user state to `%LOCALAPPDATA%` only if SharePoint folder portability is no longer a requirement.

## 15. Safe change procedure

Before editing:

1. Read this file and `CURRENT_ARCHITECTURE.md`.
2. Run the 171-test suite.
3. Identify whether the change affects Windows COM, Tkinter threading, metric aggregation, or persisted mappings.

After editing:

1. Run `python -m compileall -q app developer/tests`.
2. Run `python -m pytest developer/tests -q`.
3. Run the relevant manual checks in `TESTING_AND_RELEASE.md`.
4. Update `CHANGELOG.md` and any affected documentation.
5. Do not claim Office integration is verified unless it was run on Windows with Office.
