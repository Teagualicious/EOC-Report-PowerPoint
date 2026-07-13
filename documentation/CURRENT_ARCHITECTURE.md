# Current Architecture

This document describes the implementation in the reorganized July 10, 2026 build. The similarly named file in `archive/` is an old refactor specification and is not current.

## System overview

The application is a local Tkinter desktop workflow with five major layers:

```text
Tkinter UI
  -> platform configuration and client assignment
  -> parsers and normalization pipeline
  -> Excel report generation and optional VBA injection
  -> KPI/catalog/query services
  -> PowerPoint mapping, static fill, and optional COM live preview
```

## Repository structure

```text
app/
|-- main.py                 Startup, logging, crash hooks, DPI, event loop
|-- config/                 Paths, naming, settings, themes, logging
|-- parsers/                CSV, Excel, HTML, metric dictionary
|-- engine/                 Headless business and document-generation logic
|   `-- vba_src/            Excel search VBA source
|-- ui/                     Main workflow, settings, assignment, review
|-- mapper/                 PowerPoint mapper interface
`-- resources/              Metric dictionary and brand assets

input/                          Optional source-file staging
output/                         Default report output

workspace/
|-- templates/              Saved PPTX files and images
|-- mappings/               Platform and PPTX JSON mappings
|-- logs/                   Rotating logs
`-- settings.json           Generated user preferences

developer/
|-- tests/                  Pytest suite
|-- build/                  PyInstaller spec
`-- tools/                  Portable dependency installer
```

## Module responsibilities

### Configuration

- `config.paths` defines every resource and writable path and non-destructively migrates older layouts into the current root/user-state folders.
- `config.settings` atomically reads/writes settings and platform mappings.
- `config.naming` sanitizes user-controlled path components and provides collision-safe names.
- `config.themes` exposes light and dark theme dictionaries.
- `config.logging_setup` configures one rotating application log.

### Parsing and normalization

- `parsers.dictionary` loads aliases, breakdown levels, date patterns, and aggregation rules.
- `parsers.csv_parser` handles encoding/BOM detection, header mapping, and row conversion.
- `parsers.excel_parser` streams XLSX/XLSM worksheets and bounds header detection.
- `parsers.html_parser` scans metric-only dashboards and table-based reports and stores structure fingerprints.
- `engine.data_pipeline` applies saved platform column roles and filters parsed data by campaign.

### Excel

- `engine.excel_utils` converts parsed dictionaries into the nine-column Unified Data schema, normalizes values, collapses duplicate keys, and enforces replacement semantics on re-export.
- `engine.excel_search_dashboard` builds the Search sheet plus hidden index and configuration sheets.
- `engine.excel_writer` writes the workbook and optionally calls VBA injection.
- `engine.excel_vba` injects VBA and ActiveX controls through Excel COM.
- `engine/vba_src/modSearch.bas` contains the interactive search engine.

### Metrics and queries

- `engine.kpi` calculates the review-screen KPIs and flags suspicious data.
- `engine.metrics_catalog` builds the values exposed to PowerPoint mappings.
- `engine.query_resolver` resolves advanced metric, filter, breakdown, aggregation, top-N, and output queries.

### PowerPoint

- `engine.pptx_mapper` scans templates and persists mappings.
- `engine.pptx_formats` is the single formatting authority.
- `engine.pptx_fill` performs static python-pptx replacements.
- `engine.pptx_live` performs live COM preview and advanced shape/chart/table operations.
- `engine.pptx_thumbs` caches template thumbnails.
- `engine.template_bundle` imports and exports portable mapped-template ZIP files.

### User interface

- `ui.main_window` owns the application root, file list, reporting period, background parsing, and per-client export.
- `ui.platform_setup` creates platform column-role mappings.
- `ui.client_wizard` assigns campaigns to client names.
- `ui.review_view` renders KPI/campaign review and report-generation actions.
- `ui.settings_window` manages platforms, themes, output, templates, and debug logs.
- `ui.utils` contains shared modern widgets and thread-to-Tk callbacks.
- `mapper.mapper_window`, `mapper.sidebar`, `mapper.slide_view`, `mapper.query_builder`, and `mapper.format_popup` implement the template mapper.

## Dependency direction

```text
ui/ and mapper/
    depend on config/, parsers/, and engine/

engine/
    may depend on config/ and parsers/
    must not depend on ui/

parsers/
    may depend on config/ and engine.errors
    must not depend on ui/ or mapper/

config/
    must remain independent of application UI and document engines
```

## Threading model

Tkinter runs on the main thread. Parsing and Excel export use `ui.utils.run_in_background()`. Worker functions return plain Python data; Tk widgets are updated only in main-thread callbacks.

## Storage model

The project intentionally keeps user data beside the program for easy SharePoint-folder distribution. Code and static resources are under `app/`; application-managed state is under `workspace/`; source staging and default reports are in root `input/` and `output/`. The path layer supports PyInstaller resources while preserving all writable folders beside the executable.
