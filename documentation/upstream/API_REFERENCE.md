# API Reference

This reference lists the stable internal surfaces most likely to be used by tests, UI modules, or future extensions. Import modules by adding `app/` to `sys.path` or running through `app/main.py`.

## config.paths

### `ensure_dirs()`
Creates the workspace state tree plus the root `input/` and `output/` folders, then performs non-destructive legacy layout migration.

Important constants: `PROJECT_ROOT`, `APP_DIR`, `RESOURCE_DIR`, `WORKSPACE_DIR`, `DICT_PATH`, `LOGO_PATH`, `SETTINGS_PATH`, `MAPPINGS_DIR`, `TEMPLATES_DIR`, `IMAGES_DIR`, `INPUT_DIR`, `OUTPUT_DIR`, and `LOGS_DIR`.

## config.naming

### `safe_component(value, fallback="untitled", max_length=120) -> str`
Returns a Windows-safe path component.

### `storage_key(value, *, replace_dots=False) -> str`
Returns a stable lowercase key for mapping filenames.

### `safe_child(base_dir, component, *, fallback="untitled") -> str`
Returns a safe path beneath `base_dir` and rejects traversal.

### `unique_component(value, used, *, fallback="untitled", max_length=120) -> str`
Creates a collision-safe sanitized component (case-insensitive) and updates the supplied `used` set.

## config.settings

### `load_settings() -> dict`
Returns saved settings or `{"platforms": {}, "theme": "light"}`. The previous default `workspace/output/` value is translated in memory to root `output/`; custom output paths are preserved.

### `save_settings(data)`
Atomically writes `workspace/settings.json`.

### `load_platform_config(name) -> dict | None`
Loads one platform's sheet/column mapping.

### `save_platform_config(name, config)`
Atomically saves one platform mapping.

### `delete_platform_config(name)`
Removes one platform mapping.

## parsers.dictionary

### `load_dictionary() -> dict`
Loads and caches `app/resources/metric_dictionary.json`.

### `get_metric_aggregation(metric_name) -> str`
Returns `sum` or `avg`.

### `match_metric_alias(column_name, dictionary=None) -> str | None`
Matches a source column to a universal metric name.

### `classify_columns(header_names, dictionary=None) -> dict`
Classifies headers as metrics, campaign identifiers, levels, context, or skip columns.

## parsers

### `CSVParser(filepath).parse(build_outputs=True) -> dict`
Parses a CSV export.

### `ExcelParser(filepath).parse(build_outputs=True) -> dict`
Parses XLSX/XLSM exports.

### `HTMLParser(filepath, platform_name).parse() -> dict`
Parses HTML/HTM dashboards and tables.

## engine.data_pipeline

### `scan_file_structure(filepath) -> list[dict]`
Returns detected sheets/tables for platform setup.

### `apply_platform_config(parsed_data, platform_config) -> None`
Mutates parsed data in place using selected column roles.

### `filter_data_by_campaigns(all_parsed, campaign_names) -> list[dict]`
Returns copies containing only matching campaign rows.

## engine.excel_utils

### `SCHEMA_COLUMNS`

```python
[
    "client", "campaign", "campaign_type", "source",
    "metric_level", "metric_name", "metric_value",
    "start_date", "end_date"
]
```

### `_collect_rows(parsed_data_list, platform_name="") -> list[dict]`
Flattens parsed data into Unified Data rows.

### `_build_dataframe(new_rows, existing_rows=None) -> pandas.DataFrame`
Collapses duplicate rows and applies re-export replacement semantics.

## engine.excel_writer

### `write_to_excel(parsed_data_list, output_path, platform_name="", inject_vba=True) -> tuple[str, str | None]`
Writes Search, Unified Data, `_SearchIndex`, and `_Config`. Returns `(final_path, vba_error)`. `final_path` may be XLSM after successful injection.

## engine.kpi

### `compute_kpis(client_data) -> tuple[dict, dict, list]`
Returns `(kpi_totals, campaign_details, flags)` using best-source-per-campaign aggregation.

## engine.metrics_catalog

### `get_available_metrics(client_data, client_name="", start_date="", end_date="") -> tuple[dict, dict]`
Returns:

- a flat metric dictionary used by template filling;
- a structured dictionary with `special`, `totals`, and `breakdowns` for the mapper.

Special keys include `__client_name__`, `__date_range__`, `__start_date__`, `__end_date__`, `__start_month__`, `__end_month__`, and `__year__`.

## engine.query_resolver

### `build_simple_options(structured_metrics) -> list[dict]`
Builds mapper sidebar options.

### `resolve_query(query, client_data, client_name="", start_date="", end_date="")`
Resolves a simple key or advanced query dictionary.

### `get_available_breakdowns(client_data) -> dict`
Returns available breakdown values grouped by source type.

Advanced query example:

```json
{
  "metric": "Impressions",
  "breakdown": "device",
  "filter": "Roku",
  "agg": "sum"
}
```

Supported `agg` values: `sum` (default), `avg`/`mean`, `max`, `min`, `count`, `first`.

## engine.pptx_mapper

### `scan_template(template_path) -> list`
Returns slide and shape metadata.

### `load_template_mapping(template_filename) -> dict | None`
Loads a saved mapping.

### `save_template_mapping(template_filename, mapping)`
Atomically persists a mapping and records `_template_filename`.

### `list_available_templates() -> list[dict]`
Lists templates and mapping status.

## mapper.mapping_model

### `MappingModel(mapping=None)`
Single owner of template-mapping state for the mapper. `data` holds the canonical mapping dict in the exact persisted schema (legacy shape-level single assignments included), so `to_dict()` feeds `save_template_mapping()` and `engine.pptx_fill` unchanged.

Reads: `slide_map(slide_num)`, `shape_map(slide_num, shape_id)`, `assignments(slide_num, shape_id)` (normalized list view), `metric_formats` / `metric_format_details` dicts.

Mutations (each notifies subscribers):

- `assign_metric(slide_num, shape_id, metric, fmt="text", replace_text="", format_details=None, query=None, confirm_replace=False)` — returns `CREATED`, `APPENDED`, `UPDATED` (same metric updated in place), `NEEDS_CONFIRM` (full-text assign would replace existing assignments; nothing changed until re-called with `confirm_replace=True`), or `REPLACED`.
- `assign_image(slide_num, shape_id, image_path, image_path_abs="")`
- `set_skip(slide_num, shape_id, skip)` — replaces the shape mapping (preserved pre-model behavior: discards assignments).
- `clear_shape(slide_num, shape_id)`
- `set_metric_format(metric, fmt, details=None) -> int` — stores per-metric preferences and propagates details into every existing assignment of that metric.

`subscribe(callback)` registers `callback(event, slide_num=None, shape_id=None, metric=None)`; events are `assign`, `image`, `skip`, `clear`, `format`. `note_rendered(slide_num, shape_id, metric, display)` records live-preview bookkeeping without notifying.

## engine.pptx_fill

### `fill_template(template_path, output_path, mapping, metric_values) -> str`
Performs static text, image, query, and formatting replacement with python-pptx. Delegates to `fill_template_report()` and returns only the path.

### `fill_template_report(template_path, output_path, mapping, metric_values) -> tuple[str, FillReport]`
Like `fill_template`, but also returns a `FillReport` describing what was filled, what was missing, and what failed — including `replace_text` placeholders that no longer exist in the template.

## engine.fill_report

### `FillReport(template_path="", output_path="")`
Mutable, JSON-friendly per-fill outcome record. Fields: `template`, `output`, `timestamp`, counters `filled`, `images_filled`, `skipped`, `out_of_range`, and lists `missing_metrics`, `unmatched_placeholders`, `missing_images`, `failed_queries`, `errors`. `ok` is True when nothing the user mapped was left unfilled; `summary()` returns the post-fill dialog text; `to_dict()` returns the JSON form.

### `append_fill_history(report) -> str | None`
Appends one JSON line per fill to `workspace/logs/fill_history.jsonl`. Best-effort: failures are logged and swallowed.

## engine.pptx_formats

### `_coerce_number(value)`
Converts NumPy scalars and numeric-like values into native numbers when safe.

### `format_with_details(value, details) -> str`
Applies detailed number/date formatting.

### `_format_value(value, fmt, existing_text=None) -> str`
Applies standard text, number, percentage, currency, and date formatting with case/style matching.

## engine.pptx_live

### `PPTXLivePreview(template_path)`
Windows COM live-preview controller.

Important methods:

- `is_active()`
- `get_health()`
- `go_to_slide(slide_num)`
- `export_slide_image(slide_num, output_path=None, width=800)`
- `update_shape_text(slide_num, shape_index, new_text, replace_portion=None, expected_name=None)`
- `restore_shape_text(slide_num, shape_index)`
- `replace_shape_with_image(slide_num, shape_index, image_path)`
- `update_chart_data(slide_num, shape_index, data_dict)`
- `update_table_data(slide_num, shape_index, data_dict)`
- `save_as(output_path)`
- `cleanup()`

Public COM methods are health-tracked: after `MAX_CONSECUTIVE_FAILURES` (3) consecutive failures the preview disables itself, fires the optional `on_disabled` callback once, and static fill proceeds through python-pptx. `get_health()` returns call/failure counters, the last failed operation, the disabled reason, and `active`.

## engine.template_bundle

### `export_template_bundle(template_filename, dest_zip) -> int`
Writes a portable ZIP and returns the number of bundled images.

### `import_template_bundle(zip_path, overwrite=False) -> tuple[str, int]`
Validates and installs a bundle, returning `(template_name, image_count)`.

## Template mapping format

```json
{
  "_template_filename": "Monthly_Report.pptx",
  "slides": {
    "1": {
      "shape_mappings": {
        "2": {
          "skip": false,
          "assignments": [
            {
              "metric": "__client_name__",
              "format": "text",
              "replace_text": "CLIENT NAME"
            },
            {
              "metric": "Total Impressions",
              "format": "number",
              "replace_text": "XX",
              "format_details": {
                "format": "number",
                "decimals": 0,
                "commas": true,
                "prefix": "",
                "suffix": ""
              }
            }
          ]
        },
        "5": {
          "skip": false,
          "image_path": "workspace/templates/images/logo.png",
          "image_path_abs": "C:/optional/current/session/logo.png",
          "assignments": []
        }
      }
    }
  }
}
```

Notes:

- An assignment may also carry a `"query"` dictionary (see `engine.query_resolver`); `fill_template` resolves it when the metric key is not already in `metric_values`.
- `image_path` is stored relative to the project root for portability; `image_path_abs` keeps the absolute path for the current session.
- Older mappings may store a single assignment's fields (`metric`, `format`, `replace_text`, `format_details`, `query`) directly on the shape entry instead of in `assignments`; readers still honor that form.
