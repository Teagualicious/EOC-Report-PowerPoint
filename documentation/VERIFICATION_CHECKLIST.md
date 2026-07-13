# Manual Verification Checklist

Use this checklist for focused UI verification. The broader release checklist is in `TESTING_AND_RELEASE.md`.

## Main workflow

- [ ] Launch with `Start Ingestion Engine.bat`.
- [ ] Main window shows the branded header, file card, reporting dates, and workflow progress.
- [ ] Light and dark themes render legibly.
- [ ] Add several CSV/XLSX/XLSM/HTML files and assign platforms.
- [ ] Attempt to add an XLS file and confirm the conversion message is clear.
- [ ] Choose valid start/end dates; reject an invalid range.
- [ ] Continue; confirm the UI stays responsive while parsing.
- [ ] Client wizard opens and campaign filtering preserves selections.
- [ ] Client campaign rows span the available width and Next/Back remain in the bottom action bar.
- [ ] Assign at least two clients and confirm separate output folders in `output/` or the configured destination.
- [ ] Review KPI cards, campaign details, and data warnings.

## Settings

- [ ] Add, edit, and delete a platform mapping.
- [ ] Change theme and restart; confirm it persists.
- [ ] Change the default output folder and confirm exports use it.
- [ ] Templates tab lists templates and mapping status.
- [ ] Logs tab shows the tail of `workspace/logs/ingestion_engine.log`.
- [ ] At 125-175% display scaling, all tab labels and header subtitles remain visible.
- [ ] Platform Setup keeps Save Platform and Cancel below the scrollable mapping table.

## Excel

- [ ] Workbook contains Search, Unified Data, `_SearchIndex`, and `_Config`.
- [ ] Unified Data has the expected nine columns, filters, banding, and frozen header.
- [ ] With Excel Trust Center access enabled, export upgrades to XLSM.
- [ ] Search box is positioned correctly and does not cover the banner.
- [ ] Suggestions, date filtering, result grouping, and Copy work.
- [ ] Re-running the same period replaces matching values rather than doubling them.
- [ ] Exporting while the workbook is open gives a friendly close-file message.

## PowerPoint mapper

- [ ] Live Preview reports active when PowerPoint is available.
- [ ] Full-shape replacement works.
- [ ] Partial replacement preserves surrounding text and formatting.
- [ ] Multiple assignments in one shape work.
- [ ] Number/currency/percentage/date details match preview and final fill.
- [ ] Clear restores multi-paragraph text without merging paragraphs.
- [ ] Image replacement preserves shape geometry.
- [ ] Advanced Query can create value, table, and chart assignments.
- [ ] Save Mapping persists after restart.
- [ ] Auto-fill works for a different client and period.
- [ ] Template selector shows either a slide thumbnail or text fallback and keeps all three action buttons visible.

## Error paths

- [ ] Corrupt input identifies the file and allows valid files to continue.
- [ ] Missing platform configuration produces an actionable warning.
- [ ] PowerPoint closing unexpectedly does not crash the full application.
- [ ] Unexpected errors write a traceback to the workspace log.

## Folder migration

- [ ] Place a legacy root-level `settings.json`, `templates/`, or `mappings/` item in a test copy.
- [ ] Launch and confirm legacy settings/templates/mappings are copied into `workspace/`.
- [ ] Confirm files from `input_files/` or the previous `workspace/input/` and `workspace/output/` folders are copied into root `input/` and `output/`.
- [ ] Confirm existing destination files are not overwritten and legacy files remain in place.
