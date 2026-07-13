# Demonstration Script

## Preparation

- Run `Start Ingestion Engine.bat` once so dependencies are installed.
- Configure at least one platform.
- Place representative exports in `input/` or another known folder.
- Clear demonstration-only files from `output/`.
- Confirm the stand-in PowerPoint template is mapped.
- On Windows, enable Excel Trust Center access to the VBA project object model.

## Ten-minute demonstration

### 1. Introduce the folder and privacy model

Show the clean root layout: source exports in `input/`, generated reports in `output/`, program code in `app/`, application state in `workspace/`, documentation in `documentation/`, automated tests in `tests/`, and developer material in `developer/`. Explain that campaign data stays local.

### 2. Launch and add files

Open `Start Ingestion Engine.bat`, add mixed platform exports, and confirm each platform assignment.

### 3. Set the reporting period

Choose start and end dates and continue. Point out that parsing occurs without freezing the interface.

### 4. Assign campaigns

Enter a client name, filter the campaign list, select campaigns, clear the filter, and show that selections remain. Add a second client if time permits.

### 5. Review and inspect output

Show KPI cards and campaign details. Open the generated client folder and the Unified Data sheet.

### 6. Demonstrate Excel Search

Use a blank search, a level value, a metric, and a grouping term. Demonstrate date filtering and Copy.

### 7. Auto-fill PowerPoint

Choose a mapped template and generate the report. Show client name, reporting period, metrics, and formatting in the result.

### 8. Show one mapper assignment

Open the mapper, select a metric, assign it to a placeholder, and show the live/static preview. Mention that the mapping is reusable for every client and month.

### 9. Close with handoff and support

Show `documentation/USER_GUIDE.md` and `workspace/logs/ingestion_engine.log`. Explain that a mapped template can be exported as a portable ZIP for teammates.

## Fallbacks

| Problem | Fallback |
|---|---|
| App does not launch | Run `python app/main.py` from a terminal and inspect `workspace/logs/ingestion_engine.log`. |
| Excel VBA injection is blocked | Open the generated XLSX and demonstrate Unified Data; explain the required Trust Center setting. |
| PowerPoint live preview is unavailable or fails mid-demo | Continue in static mapper mode; after repeated preview errors the app shows one "Live Preview Off" notice and Save & Fill uses the built-in fill engine automatically. |
| Source export fails | Use a known-good sample and show the per-file error handling. |
