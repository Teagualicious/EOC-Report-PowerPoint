> **ARCHIVED - NOT AUTHORITATIVE.** This file is preserved as superseded refactor specification. Use `../MODEL_HANDOFF.md` and `../CURRENT_ARCHITECTURE.md` for the current implementation.

# Ingestion Engine — Architecture Refactor Specification

## Overview
The Spectrum Reach Reporting Ingestion Engine is a Python/tkinter desktop application that:
1. Ingests campaign data from multiple platforms (CSV, XLSX, HTML)
2. Normalizes data into a unified schema using a metric dictionary
3. Exports one workbook per client: interactive Search dashboard (VBA, injected at export) + Unified Data
4. Fills PowerPoint templates with data via COM automation
5. Provides a visual query builder for data exploration

## Excel Search Dashboard (v16.9+)
- `engine/excel_search_dashboard.py` — openpyxl side: Search sheet layout,
  hidden `_SearchIndex` (term -> kind -> canonical) and `_Config`
  (metric order/enabled/aggregation) sheets, defined names.
- `engine/vba_src/*.bas` — VBA sources (modSearch: parser, aggregator
  with best-source-per-campaign collapse, suggestion chips, copy).
- `engine/excel_vba.py` — COM injector: adds the VBA, an ActiveX
  search textbox (per-keystroke suggestions), and a runtime-built
  calendar UserForm; saves as .xlsm. Requires Excel's "Trust access to
  the VBA project object model"; falls back to plain .xlsx gracefully.
- Client Report and the old Campaign Dashboard sheet are REMOVED, as
  are the export checkboxes and Dashboard settings tab.

## Current State
Everything is in `ingestion_engine.py` (~3000 lines) plus these already-separated modules:
- `excel_writer.py` — pandas-powered Excel output
- `pptx_mapper.py` — PowerPoint template scanning, mapping, filling
- `pptx_live.py` — COM automation for live PowerPoint preview
- `query_resolver.py` — pandas query engine
- `parsers/dictionary.py` — metric alias matching
- `parsers/csv_parser.py` — CSV file parser
- `parsers/excel_parser.py` — Excel file parser
- `parsers/html_parser.py` — HTML file parser

## Target Module Structure

```
ingestion_engine/
│
├── main.py                      # Entry point only: creates IngestionEngine, calls run()
│
├── config/
│   ├── __init__.py
│   ├── settings.py              # load_settings(), save_settings(), SETTINGS_PATH
│   └── themes.py                # THEMES dict, get_theme()
│
├── ui/
│   ├── __init__.py
│   ├── main_window.py           # IngestionEngine class — main window, file import, date range, Quick Run
│   ├── settings_window.py       # SettingsWindow class — Platforms, Appearance, Dashboard Metrics, Export tabs
│   ├── platform_setup.py        # PlatformSetupWindow class — import sample file, assign column roles
│   ├── client_wizard.py         # ClientWizard class — campaign assignment wizard
│   ├── review_view.py           # Review rendering logic (_show_review_in_main, KPI calculation)
│   └── utils.py                 # _bind_mousewheel(), shared UI helpers
│
├── mapper/
│   ├── __init__.py
│   ├── mapper_window.py         # PPTXWizard class — template mapper with embedded preview
│   ├── query_builder.py         # _show_query_builder() — advanced query popup
│   └── format_popup.py          # _show_format_popup() — right-click number format
│
├── engine/
│   ├── __init__.py
│   ├── data_pipeline.py         # apply_platform_config(), filter_data_by_campaigns()
│   ├── excel_writer.py          # (already exists) write_to_excel()
│   ├── pptx_mapper.py           # (already exists) scan_template(), fill_template(), etc.
│   ├── pptx_live.py             # (already exists) PPTXLivePreview class
│   └── query_resolver.py        # (already exists) resolve_query(), build_simple_options()
│
├── parsers/
│   ├── __init__.py
│   ├── dictionary.py            # (already exists) metric alias matching
│   ├── csv_parser.py            # (already exists)
│   ├── excel_parser.py          # (already exists)
│   └── html_parser.py           # (already exists)
│
├── data/
│   ├── metric_dictionary.json   # Metric aliases, level definitions, date formats
│   └── settings.json            # User settings (auto-generated)
│
├── templates/                   # Internally stored PowerPoint templates
│   └── images/                  # Internally stored images (logos, maps)
│
├── mappings/                    # Saved template-to-metric mappings (JSON)
│
├── output/                      # Default output folder
│
├── logo.png                     # Spectrum Reach logo
├── launch.bat                   # Windows launcher
└── requirements.txt             # Python dependencies
```

## Refactoring Rules

### DO:
- Split by class — each class gets its own file
- Keep ALL existing logic intact — no rewriting behavior
- Update imports to reference new module paths
- Keep the same public API for each class (same __init__ params, same methods)
- Move constants (THEMES, KPI_METRICS, KPI_ALIASES, SCHEMA_COLUMNS) to appropriate modules
- Keep `metric_dictionary.json` path resolution working from any module

### DON'T:
- Don't change any business logic
- Don't rename classes or methods
- Don't change the data flow or function signatures
- Don't remove any features
- Don't change the UI layout or behavior
- Don't modify the parsers or engine modules (they're already clean)

## Key Dependencies Between Modules

```
main_window.py
  ├── imports: settings.py, themes.py, platform_setup.py, settings_window.py
  ├── imports: client_wizard.py, review_view.py, data_pipeline.py
  ├── imports: excel_writer.py (for export)
  └── creates: PPTXWizard (from mapper_window.py) via review_view

client_wizard.py
  ├── imports: settings.py (for export options)
  └── calls: on_complete callback → main_window._on_wizard_complete

review_view.py
  ├── imports: pandas, KPI_METRICS, KPI_ALIASES
  ├── renders into: main_window's root frame
  └── creates: PPTXWizard (from mapper_window.py)

mapper_window.py
  ├── imports: pptx_mapper.py, pptx_live.py, query_resolver.py
  ├── imports: query_builder.py, format_popup.py
  └── uses: PIL/Pillow for thumbnail rendering

query_builder.py
  ├── imports: pandas, query_resolver.py
  └── returns: query dict + selected metric to mapper_window

format_popup.py
  └── modifies: mapper_window._metric_formats dict
```

## Shared State
These variables are shared across modules and need careful handling:
- `self.t` (theme dict) — passed to every UI class
- `self.platforms` — platform configurations, stored in settings
- `self.parsed_data` — list of parsed file data dicts
- `self.available_metrics` — dict of metric_key → value
- `self._metric_formats` — dict of metric_key → format string
- `self._metric_format_details` — dict of metric_key → format config
- `self.mapping` — template mapping dict (slides → shapes → assignments)
- `self.live_preview` — PPTXLivePreview instance

## Constants to Extract

### config/themes.py
```python
THEMES = {
    "light": { ... },
    "dark": { ... },
}
def get_theme(name): ...
```

### In review_view.py or a constants file
```python
KPI_METRICS = ["Impressions", "Clicks", "Completions", "Cost"]
KPI_ALIASES = {
    "100% Completions": "Completions",
    "100% Complete": "Completions",
    "Contributions": "Completions",
    "Contribution": "Completions",
}
```

## Entry Point (main.py)
```python
from ui.main_window import IngestionEngine

if __name__ == "__main__":
    IngestionEngine().run()
```

## Testing After Refactor
Run through this workflow to verify nothing broke:
1. Launch → main window appears with logo, file import, date range
2. Settings → all 4 tabs work (Platforms, Appearance, Dashboard Metrics, Export)
3. Add Platform → import sample file, assign columns, save
4. Add files → files appear with platform dropdown
5. Set dates → calendar popup works
6. Quick Run → client assignment wizard opens maximized
7. Assign client → review appears in main window with KPIs
8. Generate Report → template selector or mapper opens
9. Mapper → slide preview, sidebar with right-click format, shape assignment
10. Save & Fill → PowerPoint generated correctly
