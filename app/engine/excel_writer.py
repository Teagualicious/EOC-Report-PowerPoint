"""Excel report writer — pandas data pipeline + openpyxl styling.

Data flow:
1. Collect raw rows from parsed data into a pandas DataFrame (excel_utils)
2. Deduplicate by composite key, aggregating numeric values
3. Pivot for dashboard sheets
4. Write to Excel with openpyxl formatting
"""

import logging
import os

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from engine.excel_utils import (
    SCHEMA_COLUMNS, _format_cell, _collect_rows, _build_dataframe,
    # re-exported for tests (pyflakes flags these; intentional):
    _normalize_value, _normalize_text,  # noqa: F401
)

logger = logging.getLogger(__name__)

# ── Styling ──
NAVY = "1B2A4A"
LIGHT_BLUE = "E8EEF4"
MED_BLUE = "C5D5E8"
LIGHT_GRAY = "F7F7F7"
MED_GRAY = "E0E0E0"
DARK_TEXT = "1A1A1A"
MUTED_TEXT = "666666"

HEADER_FILL = PatternFill("solid", fgColor=NAVY)
HEADER_FONT = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
DATA_FONT = Font(name="Calibri", size=10, color=DARK_TEXT)
SECTION_FILL = PatternFill("solid", fgColor=LIGHT_BLUE)
SECTION_FONT = Font(name="Calibri", bold=True, size=12, color=NAVY)
CAMPAIGN_HDR_FILL = PatternFill("solid", fgColor=MED_BLUE)
CAMPAIGN_HDR_FONT = Font(name="Calibri", bold=True, size=10, color=NAVY)
TOTAL_FILL = PatternFill("solid", fgColor=MED_BLUE)
TOTAL_FONT = Font(name="Calibri", bold=True, size=10, color=DARK_TEXT)
ALT_ROW_FILL = PatternFill("solid", fgColor=LIGHT_GRAY)
TITLE_FONT = Font(name="Calibri", bold=True, size=18, color=NAVY)
SUBTITLE_FONT = Font(name="Calibri", size=11, color=MUTED_TEXT)
THIN_BORDER = Border(bottom=Side(style="thin", color=MED_GRAY),
                      right=Side(style="thin", color=MED_GRAY))
CR_HEADER_FILL = PatternFill("solid", fgColor=NAVY)
CR_TITLE_FONT = Font(name="Calibri", bold=True, size=22, color="FFFFFF")
CR_DATE_FONT = Font(name="Calibri", size=12, color="B0C4DE")

MANAGED_SHEETS = ["Unified Data", "Search", "_SearchIndex", "_Config",
                  "Campaign Dashboard", "Client Report"]  # legacy names cleaned on re-export




def write_to_excel(parsed_data_list, output_path, platform_name="",
                   inject_vba=True):
    """Write the unified workbook: Search dashboard + Unified Data.

    Attempts to inject the interactive VBA search engine and save as .xlsm
    (Windows + Excel + trust setting required); otherwise leaves a plain
    .xlsx with the data intact. Returns (final_path, vba_error_or_None).
    """
    # Re-export continuity: a previous export may have been upgraded to
    # .xlsm — keep appending to it rather than forking a new .xlsx
    base, ext = os.path.splitext(output_path)
    if os.path.exists(base + ".xlsm") and not output_path.lower().endswith(".xlsm"):
        # xlsm is the live report; a same-name xlsx beside it is a stale
        # leftover from a locked delete — self-heal and use the xlsm
        if os.path.exists(output_path):
            try:
                os.remove(output_path)
            except OSError:
                logger.debug("Stale xlsx still locked", exc_info=True)
        output_path = base + ".xlsm"
    is_xlsm = output_path.lower().endswith(".xlsm")

    # Read existing data if file exists
    existing_rows = []
    wb = None
    if os.path.exists(output_path):
        try:
            wb = load_workbook(output_path, keep_vba=is_xlsm)
        except Exception:
            logger.warning("Existing report at %s is unreadable — rebuilding "
                           "it from scratch", output_path, exc_info=True)
            wb = None
    if wb is not None:
        # On .xlsm re-export keep the Search sheet — deleting it would strand
        # its VBA code-behind; only its hidden data sheets are rebuilt.
        keep = {"Search"} if is_xlsm else set()
        if "Unified Data" in wb.sheetnames:
            ws_old = wb["Unified Data"]
            if ws_old.max_row >= 2:
                hdrs = [ws_old.cell(row=1, column=c).value for c in range(1, ws_old.max_column + 1)
                        if ws_old.cell(row=1, column=c).value]
                for r in range(2, ws_old.max_row + 1):
                    rd = {h: ws_old.cell(row=r, column=i+1).value for i, h in enumerate(hdrs)}
                    if rd.get("metric_name"):
                        existing_rows.append(rd)
        for name in MANAGED_SHEETS:
            if name in wb.sheetnames and name not in keep:
                del wb[name]
    if wb is None:
        wb = Workbook()
        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]

    # Collect and deduplicate
    new_rows = _collect_rows(parsed_data_list, platform_name)
    df = _build_dataframe(new_rows, existing_rows)

    if df.empty:
        raise ValueError(
            "No data to export — no rows matched the selected campaigns "
            "and date range.")

    # ── Unified Data sheet ──
    ws = wb.create_sheet("Unified Data")
    for ci, h in enumerate(SCHEMA_COLUMNS, 1):
        c = ws.cell(row=1, column=ci, value=h)
        c.font = HEADER_FONT; c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center"); c.border = THIN_BORDER

    for ri, (_, rd) in enumerate(df.iterrows(), 2):
        for ci, col in enumerate(SCHEMA_COLUMNS, 1):
            val = rd.get(col)
            if pd.isna(val): val = None
            c = ws.cell(row=ri, column=ci, value=val)
            c.font = DATA_FONT; c.border = THIN_BORDER
            if col == "metric_value" and val is not None:
                _format_cell(c, val)

    widths = {"client": 18, "campaign": 35, "campaign_type": 15, "source": 18,
              "metric_level": 25, "metric_name": 20, "metric_value": 15,
              "start_date": 14, "end_date": 14}
    for ci, col in enumerate(SCHEMA_COLUMNS, 1):
        ws.column_dimensions[get_column_letter(ci)].width = widths.get(col, 15)
    last_col = get_column_letter(len(SCHEMA_COLUMNS))
    # Pure-openpyxl polish: real Excel Table (filter buttons + banding),
    # frozen header, brand tab. A Table replaces auto_filter.
    from openpyxl.worksheet.table import Table, TableStyleInfo
    tbl = Table(displayName="UnifiedData", ref=f"A1:{last_col}{len(df) + 1}")
    tbl.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2",
                                        showRowStripes=True)
    ws.add_table(tbl)
    ws.freeze_panes = "A2"
    ws.sheet_properties.tabColor = "003057"

    # ── Search dashboard (sheet 0) + hidden index/config ──
    from engine.excel_search_dashboard import build_search_dashboard
    build_search_dashboard(wb, df)

    try:
        wb.save(output_path)
    except PermissionError as e:
        raise PermissionError(
            f"Cannot write {output_path} — the file is open in Excel. "
            "Close it and try again.") from e
    logger.info("Excel report written: %s (%d data rows)", output_path, len(df))

    # ── Interactive search: inject VBA and upgrade to .xlsm ──
    vba_error = None
    if inject_vba and not is_xlsm:
        from engine.excel_vba import inject_search_vba
        output_path, vba_error = inject_search_vba(output_path)
        if vba_error:
            logger.warning("Search VBA not injected: %s", vba_error)
    return output_path, vba_error
