# Spectrum Reach Reporting Ingestion Engine

## User Guide

**Updated:** July 13, 2026

## 1. What the application does

The Reporting Ingestion Engine reads raw campaign exports, normalizes metrics, assigns campaigns to clients, creates a searchable Excel workbook for each client, and fills mapped PowerPoint templates.

Supported source formats:

- CSV
- XLSX
- XLSM
- HTML and HTM

Legacy XLS files are not supported directly. Open them in Excel and save them as XLSX first.

## 2. Requirements

- Windows 10 or Windows 11
- Python 3.10 or newer
- Microsoft Excel for interactive VBA search injection
- Microsoft PowerPoint for live mapper preview and COM features

The rest of the application can still parse files, review data, write a normal XLSX, and statically fill many PowerPoint templates without Office COM.

## 3. First launch

1. Copy the complete `Jughead-Data-Engine` folder to the computer.
2. Double-click **`Start Ingestion Engine.bat`**.
3. The launcher finds Python, checks the version, and installs missing packages from `app/requirements.txt` on the first run. If the folder contains a `wheelhouse/` directory (included in the portable release zip, `IngestionEngine-<version>-portable-win64.zip`), the packages install from it **offline** — no network needed. Otherwise they install from the internet as before. The portable zip's bundled packages target Python 3.12; on another Python version the launcher automatically falls back to the internet install.
4. The main window opens.

When support is needed, send `workspace/logs/ingestion_engine.log`.

## 4. One-time platform setup

A platform configuration tells the application which columns in a particular export represent campaigns, metrics, and breakdown dimensions.

1. Open **Settings**.
2. Select **Platforms** and choose **Add Platform**.
3. Enter a platform name.
4. Import a representative sample CSV, XLSX, XLSM, HTML, or HTM file.
5. Review the detected sheets/tables and each column role:
   - `metric`
   - `campaign_id`
   - `level: device`, `level: zip`, or another breakdown
   - `skip`
6. Select only the columns that should be used.
7. Save the platform.

Platform mappings are stored in `workspace/mappings/` and are reused on later runs.

## 5. Monthly reporting workflow

### Step 1: Add files

Select all exports for the reporting period. Files can come from different platforms and can contain multiple clients.

Each file has a platform selector. Confirm the correct platform before continuing.

### Step 2: Set the date range

Choose the reporting start and end dates. These dates are written into the Excel data and exposed to PowerPoint template mappings.

### Step 3: Continue and parse

The application parses files in the background. If one file fails, the error dialog identifies it and allows the remaining readable files to continue.

### Step 4: Assign campaigns to clients

For each client:

1. Enter the client name.
2. Search or filter the campaign list when necessary.
3. Select the campaigns belonging to that client.
4. Continue until every required client is assigned.

Filtering the campaign list does not remove prior selections.

### Step 5: Review

The review screen shows:

- total Impressions
- total Clicks
- total Completions
- total Cost
- per-campaign details
- warnings for suspicious or zero-value data

The workbook is generated before the review screen appears.

### Step 6: Generate the PowerPoint report

Choose a mapped template and auto-fill it, or open the mapper to create/edit a mapping. The generated deck is saved in the same client output folder as the workbook.

## 6. Output files

By default, reports are written to:

```text
output/<Client Name>/
```

The workbook name is:

```text
<Client Name>_unified_report.xlsx
```

When Excel VBA injection succeeds, the workbook becomes XLSM. If injection is unavailable, the XLSX remains usable and a warning explains why the interactive search was not added.

The default output folder can be changed in **Settings > Export**.

## 7. Excel Search dashboard

Each workbook contains:

- `Search`
- `Unified Data`
- hidden `_SearchIndex`
- hidden `_Config`

The Search sheet supports comma-separated terms. Examples:

| Search | Result |
|---|---|
| blank | Campaign totals with default metric columns |
| `28167` | Rows for that matching level value |
| `28167, impressions` | The same rows with Impressions selected |
| `network` | Group by network |
| `campaign, network, impressions` | Campaign and network labels with Impressions |

Metric terms appear as columns in the order typed. Date controls restrict the search result to a reporting period.

### Required Excel setting

For the app to inject the interactive search VBA, Excel must allow programmatic access to the VBA project:

1. Excel > File > Options > Trust Center.
2. Trust Center Settings > Macro Settings.
3. Enable **Trust access to the VBA project object model**.
4. Enable content when opening the generated XLSM.

If this setting is unavailable due to company policy, the workbook still contains Unified Data and the non-injected Search layout.

## 8. PowerPoint template mapper

The mapper displays available metrics, the current slide, and its shapes.

### Assigning text or numbers

1. Select a metric in the sidebar.
2. Select all or part of a target text box.
3. Assign the metric to the shape.

Selecting a substring creates a partial replacement and preserves surrounding text. A shape can contain multiple assignments.

### Formatting

Right-click a metric to set:

- number, currency, percentage, date, or text format
- decimal places
- comma grouping
- prefix
- suffix

The same formatting logic is used by the mapper preview and final fill.

### Dates and case

The application detects placeholder date style and text case from the specific replacement target. This allows an all-caps client placeholder to remain all caps and a date placeholder to keep its intended format.

### Images

Choose **Browse Image**, select a file, and assign it to a shape. The image replaces the shape at the same position and size. Images are copied into `workspace/templates/images/` for portability.

### Advanced queries

The Query Builder can select a metric, breakdown, campaign/value filter, aggregation, top-N rule, and output type. Rate metrics should use average rather than sum.

### Saving

- **Save Mapping** stores the mapping for later clients and periods.
- **Save & Fill Template** stores the mapping and creates the current report.

After a fill, the confirmation dialog summarizes what was filled. If anything could not be filled — a missing metric, image, or placeholder — the dialog is titled **Report Generated — With Gaps** and lists what was skipped so it can be corrected.

## 9. Template sharing

In **Settings > Templates**:

- **Export** packages a mapped PPTX, its JSON mapping, and referenced images into one ZIP bundle.
- **Import** validates and installs a teammate's bundle.

Imports reject unsafe archive paths, unsupported bundle formats, oversized archives, duplicate image names, and missing image references.

## 10. Folder guide

```text
Jughead-Data-Engine/
|-- Start Ingestion Engine.bat
|-- input/                  Optional source-file staging
|-- output/                 Default generated reports
|-- app/                    Program files; do not edit for normal use
|-- workspace/
|   |-- templates/          Saved PPTX templates and images
|   |-- mappings/           Platform and PPTX mappings
|   |-- logs/               Support logs
|   `-- settings.json       Created after saving settings
|-- documentation/
|-- tests/                  Automated tests; not needed for normal use
`-- developer/
```

To move the application to another computer, copy the entire folder. At minimum, preserve `workspace/templates/`, `workspace/mappings/`, and `workspace/settings.json`.

## 11. Troubleshooting

| Symptom | Action |
|---|---|
| Python 3 was not found | Install Python 3.10+ and select Add Python to PATH, or contact IT. |
| A platform has no configuration | Create or update it in Settings > Platforms. |
| XLS file cannot be selected | Save it as XLSX in Excel first. |
| An Excel source is invalid or corrupt | Re-download it or open and resave it in Excel. |
| Output cannot be written | Close the existing workbook in Excel and retry. |
| Interactive search was not enabled | Verify Excel is installed and the Trust Center setting is enabled. The XLSX data remains valid. |
| Mapper says static mode | PowerPoint COM is unavailable. Static filling may still work. |
| A "Live Preview Off" notice appears in the mapper | PowerPoint stopped responding, so the preview turned itself off. Keep working; Save & Fill still generates the report. Reopen the mapper to restore the preview. |
| Report dialog says "With Gaps" | Some mapped items could not be filled. The dialog lists them; check the named metrics, images, or placeholder text and fill again. |
| A total looks wrong | Review campaign detail, confirm platform column roles, and verify rate metrics are configured as averages. |
| Unexpected crash | Send `workspace/logs/ingestion_engine.log` to support. |
