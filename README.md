# Spectrum Reach Reporting Ingestion Engine

A local Windows desktop application that converts campaign exports into searchable Excel workbooks and fills mapped PowerPoint report templates.

## Development and AI handoff

- **`CLAUDE.md` / `STATUS.md`** — repository session workflow: read STATUS.md first, tests are the completion gate, synthetic fixtures only.
- **`AI_CONTEXT.md`** — concise design rules, pitfalls, and preferred change workflow for AI assistants.
- **`documentation/MODEL_HANDOFF.md`** — detailed authoritative architecture and implementation handoff.

## Start here (end users)

1. Double-click **`Start Ingestion Engine.bat`**.
2. On the first run, the launcher checks for Python 3.10+ and installs any missing packages.
3. Add platform exports, set the reporting period, assign campaigns to clients, review totals, and generate reports.

The application works with CSV, XLSX, XLSM, HTML, and HTM exports. Legacy XLS files must be saved as XLSX first.

## Start here (developers)

```bash
pip install -r requirements.txt
pytest
python app/main.py
```

GitHub Actions runs the test suite on every push and pull request.

## Folder layout

```text
.
|-- Start Ingestion Engine.bat   User launcher
|-- README.md                    This quick-start file
|-- CLAUDE.md / STATUS.md        Repo session workflow and current project state
|-- AI_CONTEXT.md                AI assistant orientation for the application code
|-- requirements.txt             Dev/CI dependencies (includes app/requirements.txt)
|-- input/                       Optional source-export staging
|-- output/                      Default generated reports
|-- app/                         Program code, dependencies, and resources
|-- workspace/                   Settings, templates, mappings, and logs
|-- documentation/               User, technical, testing, and handoff docs
|-- tests/                       pytest suite (synthetic fixtures only)
`-- developer/                   Build configuration and developer tools
```

The default generated reports are written to `output/<client>/`. A different output folder can be selected in Settings.

## Documentation

- `documentation/USER_GUIDE.md` - complete end-user workflow
- `documentation/MODEL_HANDOFF.md` - authoritative continuation guide for another developer or model
- `documentation/TECHNICAL_GUIDE.md` - data flow, storage, architecture, and integration behavior
- `documentation/CURRENT_ARCHITECTURE.md` - current module map and dependencies
- `documentation/API_REFERENCE.md` - important public functions and data contracts
- `documentation/TESTING_AND_RELEASE.md` - automated and Windows acceptance testing
- `documentation/CHANGELOG.md` - current reviewed build changes

## Current validation status

- 249 automated tests pass (run `pytest` from the repo root).
- All Python modules compile.
- The Tkinter application launches in a graphical smoke test.
- CSV and Excel large-file performance was improved in the July 10, 2026 review.
- Windows-only Excel VBA injection and PowerPoint COM automation still require final acceptance testing on a Windows machine with Microsoft Office installed (see `documentation/reviews/MAPPER_RELIABILITY_ROADMAP_2026-07-12.md` for the current checklist).

## Privacy and network use

Campaign data is processed locally. The application does not upload reports to a cloud service. An internet connection is only needed when the launcher must install Python packages for the first time. No real client or campaign data is committed to this repository — test fixtures are synthetic.
