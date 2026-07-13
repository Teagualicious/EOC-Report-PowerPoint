> **ARCHIVED - NOT AUTHORITATIVE.** This file is preserved as archived external design prompt. Use `../MODEL_HANDOFF.md` and `../CURRENT_ARCHITECTURE.md` for the current implementation.

# Spectrum Reach Reporting Ingestion Engine — Project Brief for Completion

## What This Is

A desktop reporting automation tool built for Spectrum Reach (Charter Communications) Account Executives. It replaces a manual process where AEs spend hours each month downloading campaign performance data from advertising platforms, copying numbers into Excel spreadsheets, and manually filling PowerPoint templates for client presentations.

The tool ingests raw data files (Excel exports from Architect/AudienceTrak platforms), normalizes them through a universal metric dictionary, generates branded Excel dashboards, and auto-fills PowerPoint presentation templates — turning a multi-hour manual workflow into a few clicks.

## The Business Problem

Spectrum Reach AEs manage dozens of advertising clients. Each month they must:
1. Download campaign data exports from internal platforms (Architect, AudienceTrak, MerchantService)
2. Open each export, find the relevant metrics across multiple sheets (Data by Devices, Data by Geo, Data by Creative, etc.)
3. Manually calculate totals, since exports have no campaign summary — only breakdown sheets
4. Copy numbers into a branded Excel report
5. Copy the same numbers into a PowerPoint template, replacing placeholder text
6. Repeat for every client

This takes 2-4 hours per client per month. With 20+ clients per AE, it's a massive time sink that's also error-prone.

## What We Built

A Python/tkinter desktop application with:

**Data Pipeline (pandas-powered):**
- Multi-format file ingestion (CSV, XLSX, HTML)
- Universal metric dictionary for normalizing column names across platforms
- Platform configuration system (configure once per data source, reuse forever)
- Automatic aggregation: groupby/sum handles deduplication, DMA aggregation, breakdown totals
- Smart KPI calculation: when no campaign summary sheet exists, sums from breakdown data and picks the highest total per metric across all sources

**Excel Output (openpyxl):**
- Three-sheet workbook: Client Report (client-facing), Campaign Dashboard (internal), Unified Data (raw)
- Branded with Spectrum Reach logo, configurable metrics, autofilters on all sheets
- Per-client folder export with separate files per client

**PowerPoint Automation (python-pptx + win32com):**
- Template scanning: reads all shapes, text, tables, charts from any .pptx
- Visual mapper: sidebar with metrics, center slide preview (COM-rendered thumbnail), shape list
- Click-to-assign with partial text replacement (highlight specific text to replace)
- Multi-assignment per shape (replace "Client Name" AND "DATE" in the same text box)
- Date auto-formatting: detects format from existing PowerPoint text, converts data dates to match
- Smart text formatting: CamelCase splitting, case matching to existing text style
- Image replacement for logos with aspect-ratio-aware scaling
- COM live preview: PowerPoint runs minimized, changes render as thumbnails
- Template portability: templates and images copied internally, relative paths for SharePoint distribution
- One-click auto-fill: pre-mapped templates fill instantly with new client data

**Query System:**
- Simple options: Total Impressions, Total Completions, Client Name, Date Range, etc.
- Advanced query builder with visual pivot table, multi-select filters (campaigns, breakdowns, values), and output as value/table/chart data
- Queries are reusable across clients — the template mapping stores queries, not values

## Current State

The application works end-to-end. A user can import files, assign clients, export Excel reports, and generate PowerPoint presentations. However, it was built iteratively over many sessions and needs engineering polish before it's production-ready.

### What's Working
- Full data pipeline from raw files to Excel output
- PowerPoint template mapping and filling
- Review screen with dynamic KPI cards
- Template selector for one-click report generation
- Date/number formatting systems
- Portable template storage

### What Needs Work

**Architecture:**
The main application file (ingestion_engine.py) is ~3000 lines containing all UI classes. It needs to be split into modules. See ARCHITECTURE.md for the exact target structure. The engine modules (excel_writer, pptx_mapper, pptx_live, query_resolver, parsers) are already cleanly separated.

**Bugs and Edge Cases:**
- COM text replacement sometimes causes text wrapping in narrow text boxes (auto-widen is implemented but needs testing)
- Floating point precision from pandas .sum() occasionally produces values like 1.0000000000000009 (cleanup is implemented but may miss edge cases)
- Number formatting: values should display without decimals by default for whole numbers
- Window management: some popups may lose focus or hide behind parent windows
- Dead code: old date step wizard (~150 lines) is still in the file but never called
- Error handling: many operations have bare `except: pass` that should log or show user-friendly messages

**Missing for Production:**
- Comprehensive error handling with user-facing messages
- Input validation (empty files, corrupt data, missing columns)
- Logging system for debugging
- Unit tests for the data pipeline
- .exe packaging with PyInstaller for distribution without Python installed
- User Guide PDF and Technical Documentation PDF
- Demo script for the July 14th leadership presentation

## Files Included

Read these first:
- `ARCHITECTURE.md` — target module structure and refactoring rules
- `PROJECT_STATE.md` — complete version history, technical decisions, known issues
- `API_REFERENCE.md` — internal API documentation for every module

Core application:
- `ingestion_engine.py` — main UI (~3000 lines, needs splitting)
- `excel_writer.py` — pandas-powered Excel output
- `pptx_mapper.py` — PowerPoint template operations
- `pptx_live.py` — COM automation for live preview
- `query_resolver.py` — pandas query engine
- `parsers/` — CSV, Excel, HTML parsers + metric dictionary
- `metric_dictionary.json` — aliases, level definitions, date formats

## Definition of Done

The project is complete when:

1. **Architecture is clean** — modules are focused and reasonably sized (no hard line cap — split when a file covers more than one concern), clear module boundaries, proper imports
2. **Full workflow runs without errors** — import → assign → export → review → generate PowerPoint
3. **Error handling is comprehensive** — no crashes on bad input, every error shows a message
4. **Documentation exists** — User Guide PDF, Technical Documentation PDF, inline code comments
5. **Demo script exists** — step-by-step guide for the July 14th presentation
6. **Tests exist** — at least for the data pipeline (parsing, aggregation, KPI calculation)
7. **Package is distributable** — launch.bat works on a fresh Windows machine, or .exe via PyInstaller
8. **Code is clean** — no dead code, no debug prints, consistent style, docstrings on public functions

## Technical Environment
- Python 3.12+ on Windows 10/11
- tkinter for UI (no web framework)
- pandas for data pipeline
- openpyxl for Excel
- python-pptx for PowerPoint file operations
- pywin32 (win32com) for PowerPoint COM automation (Windows only)
- Pillow for image processing
- tkcalendar for date pickers
- All machines have Microsoft Office (PowerPoint, Excel) installed

## Constraints
- Must run on corporate Windows machines behind Charter's network
- No external API calls or cloud dependencies — everything runs locally
- Must be distributable via SharePoint as a folder (no installer required)
- Templates and images must be self-contained in the application folder
- COM automation requires PowerPoint to be installed
