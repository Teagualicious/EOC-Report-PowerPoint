# Testing and Release Guide

## Automated validation

Run from the project root:

```bash
python -m pip install -r requirements.txt
python -m compileall -q app tests
python -m pytest tests -q
```

Expected result for the current build:

```text
249 passed
```

GitHub Actions (`.github/workflows/ci.yml`) runs the same suite on every push and pull request (Python 3.12, Ubuntu).

## What automated tests cover

- CSV, Excel, and HTML parsing
- BOM and encoding behavior
- metric alias and aggregation rules
- platform column-role application
- client campaign filtering
- KPI and metric catalog consistency
- per-campaign source selection
- repeat-export replacement semantics
- Unified Data and Search workbook generation
- PowerPoint formatting helpers
- query resolution
- safe names and collision handling
- template bundle import/export and malicious archive cases
- fill report summaries and `fill_history.jsonl` telemetry
- live-preview COM health tracking and python-pptx fallback
- mapping-model state ownership (assignment semantics, schema contracts, observers)
- root input/output layout migration and prior-default settings normalization
- headless parse-to-export smoke workflows

## Windows acceptance test

Use a Windows 10/11 machine with current desktop Excel and PowerPoint.

### Startup

- Double-click `Start Ingestion Engine.bat`.
- Confirm first-run package installation succeeds.
- Confirm the main window renders correctly at 100%, 125%, and 150% display scaling.
- Confirm `workspace/logs/ingestion_engine.log` is created.

### Platform setup and parsing

- Add a CSV platform and parse a normal export.
- Add an XLSX/XLSM platform with multiple sheets.
- Add an HTML platform with both a metric-only page and a table report.
- Confirm an XLS file is rejected with conversion guidance.
- Confirm a corrupt or locked file produces a readable error and other files can continue.

### Client assignment

- Select campaigns, filter the list, clear the filter, and confirm selections remain.
- Use two client names that sanitize similarly and confirm separate output folders are created.

### Excel

- Generate a workbook and verify Search, Unified Data, `_SearchIndex`, and `_Config`.
- Enable Excel Trust Center access to the VBA project object model.
- Confirm VBA injection creates XLSM, the search box appears in the intended position, suggestions work, calendar opens, Copy works, and no stale XLSX remains.
- Re-run the same period and confirm values are replaced rather than doubled.
- Run a second period and confirm both periods remain.
- Keep the workbook open during export and confirm the user receives a close-file message.

### PowerPoint

- Open the mapper with live preview active.
- Assign full-shape and partial text replacements.
- Put multiple assignments in one text box.
- Verify all-caps text and several date placeholder styles.
- Verify number, currency, percentage, decimals, commas, prefix, and suffix formatting match preview and final deck.
- Replace an image and confirm geometry is preserved.
- Save mapping, restart the application, and auto-fill with a different client.
- Test one table, one ordinary chart, and one multi-series comparison chart.
- Confirm Save & Fill and Auto-Fill dialogs show the fill summary, a mapping with gaps shows "Report Generated — With Gaps", and `workspace/logs/fill_history.jsonl` gains a line per fill.
- Close PowerPoint unexpectedly and verify a single "Live Preview Off" notice appears, the mapper stays usable, and Save & Fill produces the report via the built-in engine.
- Run the full mapper-reliability drill in `reviews/MAPPER_RELIABILITY_ROADMAP_2026-07-12.md` if any of the above checks fail or the fill/preview code has changed.

### Template bundles

- Export a mapped template with images.
- Import it into a clean copy of the application.
- Confirm the mapping and image references work on the new machine.
- Confirm overwrite prompts behave correctly.

## Release package checklist

- Automated suite passes.
- No untracked customer data is included in `input/`, `output/`, or `workspace/logs/`.
- Default templates and mappings intended for distribution are present.
- Documentation dates, test count, and known limitations are current.
- `app/requirements.txt` is installable on a clean supported Python version.
- Portable `app/vendor/` is omitted unless intentionally producing a portable package.
- Windows Office acceptance was completed or explicitly marked as outstanding.

## PyInstaller check

Build from the project root:

```bash
python -m PyInstaller developer/build/ingestion_engine.spec --noconfirm
```

Then copy required default workspace templates/mappings beside the executable and repeat the Windows acceptance test against the packaged build.
