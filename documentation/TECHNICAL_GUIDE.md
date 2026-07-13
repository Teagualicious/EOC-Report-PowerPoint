# Spectrum Reach Reporting Ingestion Engine

## Technical Guide

**Updated:** July 13, 2026

## 1. Purpose and runtime

The application is a Python 3.10+ Tkinter desktop system for campaign-export ingestion, metric normalization, Excel report generation, and PowerPoint template automation.

Primary dependencies are pandas, NumPy, openpyxl, Beautiful Soup, python-pptx, Pillow, tkcalendar, and pywin32. Package versions are pinned in `app/requirements.txt`; pywin32 carries a `; sys_platform == "win32"` marker so non-Windows installs skip it. Development and CI installs use the root `requirements.txt`, which includes `app/requirements.txt` plus pytest.

## 2. Startup lifecycle

`app/main.py`:

1. Prefers dependencies in `app/vendor/` when present.
2. Enables DPI awareness on Windows.
3. Configures rotating file and console logging.
4. Creates the workspace folder tree and imports missing files from the old root-level layout without deleting the originals.
5. Registers Python and Tkinter exception hooks.
6. Creates `ui.main_window.IngestionEngine` and starts Tk.

The Windows launcher checks Python 3.10+, probes required imports, installs dependencies only when necessary, and leaves the console open after a startup failure.

## 3. Path and storage model

`config.paths` is the only authority for application paths.

Read-only resources:

- `app/resources/metric_dictionary.json`
- `app/resources/assets/`

Writable state:

- `workspace/settings.json`
- `workspace/mappings/`
- `workspace/templates/`
- `input/`
- `output/`
- `workspace/logs/`

A custom output folder can be stored in settings. Under PyInstaller, resources are read from `sys._MEIPASS/resources`, while the workspace and the root `input/` and `output/` folders remain beside the executable.

## 4. Configuration persistence

Settings and mapping writes use a temporary file plus `os.replace()` to prevent partial JSON files after a crash.

`settings.json` normally contains:

```json
{
  "platforms": {
    "Architect": {"configured": true}
  },
  "theme": "light",
  "output_folder": "C:/Reports"
}
```

Platform mappings are separate JSON files under `workspace/mappings/`, named `platform_<key>.json` where the key comes from `config.naming.storage_key()`. This keeps large column-role definitions out of the general settings file.

## 5. Metric dictionary

`app/resources/metric_dictionary.json` defines:

- aliases for universal metric names
- additive versus average aggregation
- breakdown/level names
- context columns
- columns to skip
- supported date placeholder patterns

`parsers.dictionary` caches the dictionary in memory. Tests reset that cache between cases.

The dictionary is the source of truth for aggregation. Any new rate metric must be added with average aggregation before it is used in KPI, Excel, query, or PowerPoint totals.

## 6. Parsers

### CSV

`CSVParser` handles encoding and BOM behavior, detects headers, keeps acronyms intact, and produces campaign or breakdown rows.

### Excel

`ExcelParser` supports XLSX and XLSM. It bounds header searching and uses read-only iteration for large sheets to avoid materializing full worksheets multiple times.

### HTML

`HTMLParser` handles metric-only pages and tabular reports. It can persist structure fingerprints in `workspace/mappings/` to make repeated parsing more reliable.

All parsers return dictionaries compatible with the contract in `MODEL_HANDOFF.md`.

## 7. Platform mapping pipeline

`scan_file_structure()` reads a representative sample and returns sheet/table headers, sample values, row counts, and suggested column classifications.

`apply_platform_config()` uses the saved sheet/column roles to rebuild:

- `campaign_metrics`
- `level_data`
- flat `metrics`
- `mapped_metrics`

If no configured sheet matches, existing parser output is preserved and a warning is logged rather than silently deleting data.

## 8. Unified Data schema

Excel output uses these columns:

```text
client
campaign
campaign_type
source
metric_level
metric_name
metric_value
start_date
end_date
```

The full row key is every column except `metric_value`.

Within a new import batch, duplicate additive rows are summed and duplicate ratio rows are averaged. During re-export, a new matching full key replaces its prior value. Nonmatching rows are retained, which supports multiple reporting periods without duplicating a rerun of the same period.

## 9. KPI and metric catalog rules

Campaign exports often repeat metrics in campaign summary, device, geography, or other breakdown sources.

For additive metrics:

1. group by campaign, metric, and source;
2. sum each source group;
3. select the maximum source total per campaign and metric;
4. sum those selected campaign totals.

For ratio metrics:

- average within campaigns;
- average campaign values for client-level catalog values;
- expose keys as `Avg <Metric>` rather than `Total <Metric>`.

`engine.kpi` and `engine.metrics_catalog` intentionally follow the same source-selection rule.

## 10. Excel generation

`write_to_excel()` reads an existing report when present, captures Unified Data rows, removes managed sheets, rebuilds the current report data, and saves the workbook. On an XLSM re-export the Search sheet is kept so its VBA code-behind survives, and an export aimed at a `.xlsx` path is redirected to an existing same-name `.xlsm`.

Managed sheets:

- Search
- Unified Data
- _SearchIndex
- _Config
- old Campaign Dashboard and Client Report sheets, which are removed during rebuild

The Unified Data sheet is an Excel table with filters, row banding, widths, formats, and a frozen header.

### VBA injection

On Windows, `excel_vba.inject_search_vba()` opens Excel through COM, injects `engine/vba_src/modSearch.bas`, a runtime-built calendar UserForm with a day-button relay class, and Search-sheet event handlers, adds an ActiveX search box, and saves as XLSM.

VBA injection is recoverable. If it fails, the normal XLSX remains and the UI reports the reason.

## 11. PowerPoint mapping and fill

`pptx_mapper.scan_template()` first uses python-pptx and may fall back to COM for unsupported cases.

Mappings are stored by template filename. Text assignments can target a full shape or a substring. Multiple assignments may coexist in one shape.

`pptx_formats` centralizes:

- NumPy scalar coercion
- number/currency/percentage formatting
- detailed decimals, commas, prefixes, and suffixes
- date-format detection and rendering
- text case matching

Static fill uses python-pptx. `pptx_fill.fill_template_report()` returns `(path, report)`, where the `engine.fill_report.FillReport` records filled assignments, missing metrics, unmatched `replace_text` placeholders, missing images, and failed queries; `fill_template()` keeps its original signature. The mapper shows the report summary after every fill and appends each report as one JSON line to `workspace/logs/fill_history.jsonl`.

Live preview uses COM and a working copy, with health tracking on every COM call: after 3 consecutive failures the preview disables itself, fires a one-time `on_disabled` notice, and Save & Fill falls back to the python-pptx engine; a success resets the counter. Image replacement preserves shape geometry.

## 12. Query engine

An advanced query can specify:

```json
{
  "metric": "Impressions",
  "breakdown": "device",
  "filter": "Roku",
  "agg": "sum",
  "top_n": "10",
  "output": "value"
}
```

Supported aggregations include sum, average/mean, max, min, and count. For breakdown `all`, source selection occurs per campaign to avoid dropping campaigns that rely on different source types.

## 13. Template bundles

A portable template bundle contains:

```text
manifest.json
mapping.json
template/<template-name>.pptx
images/<referenced files>
```

The importer validates format/version, safe filenames, path traversal (absolute or `..`/`.` archive paths are rejected), total file count (500 maximum), expanded size (500 MB maximum), duplicate image names, and mapping references. Extraction is controlled and the final template write is atomic.

## 14. UI and threading

Tkinter widgets are created and changed only on the main thread. `ui.utils.run_in_background()` runs parsing and export functions in a worker thread, pushes results through a queue, and schedules success/error callbacks with `root.after()`.

UI code must not call Tk APIs from the worker function.

## 15. Error handling and logging

- Parser errors carry a concise `user_message`.
- Per-file parse failures can be skipped while valid files continue.
- Unhandled Python and Tk callback exceptions are logged with tracebacks and shown in an error dialog.
- Logs are written to `workspace/logs/ingestion_engine.log`, rotate at 1 MB, and retain three backups.
- Each PowerPoint fill appends an outcome record to `workspace/logs/fill_history.jsonl` (best-effort; a broken history file never blocks report generation).
- COM failures are logged and should degrade gracefully.

## 16. Security and filesystem safety

- User-facing folder and filename components are sanitized.
- Collision handling prevents two different names from overwriting one another after sanitization.
- Safe child-path helpers prevent directory traversal.
- Template bundle extraction never trusts raw ZIP paths.
- Atomic writes reduce corruption risk.
- Legacy migration is copy-only and never overwrites existing destination files.

## 17. Automated validation

The current suite contains 249 passing tests in `tests/` at the repo root. It covers pure data/business behavior, document generation that can be tested without Office, Windows DPI sizing helpers, PowerPoint template-preview fallbacks, fill-outcome reporting, and live-preview health tracking.

Run from the repo root:

```bash
python -m compileall -q app tests
python -m pytest tests -q
```

See `TESTING_AND_RELEASE.md` for the Windows Office acceptance pass.

## 18. Packaging

### Portable Python folder

Run `developer/tools/make_portable.bat` to install dependencies into `app/vendor/`. `app/main.py` automatically prefers that directory.

### PyInstaller

From the project root:

```bash
python -m pip install pyinstaller
python -m PyInstaller developer/build/ingestion_engine.spec --noconfirm
```

Copy the desired default workspace templates/mappings into the distribution beside the executable. Verify resource loading and Office COM on the target machine.

## 19. Known limitations

- Excel VBA and PowerPoint COM are Windows/Office dependent.
- Complex comparison charts with multiple series need broader real-deck testing.
- Static python-pptx behavior cannot perfectly reproduce every PowerPoint object type.
- A renamed source worksheet may require updating its platform mapping.
- The app intentionally stores user state beside the program for portable folder distribution rather than in the Windows profile.
