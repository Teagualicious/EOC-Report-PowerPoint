# Workspace

This folder contains application-managed state. The application creates missing subfolders automatically.

```text
workspace/
|-- templates/       Saved PowerPoint templates and template images
|-- template_store/  Template-first ingested stores (created on first ingest)
|-- mappings/        Import profiles, platform configs, and PowerPoint mappings
|-- staging/         Editable Excel staging workbooks (Stage 2)
|-- dictionary/      Generated dictionary working files (reserved)
|-- logs/            Rotating application logs
`-- settings.json    Created after settings are saved
```

The user-facing `input/` and `output/` folders are located directly in the project root folder for easier navigation; rejected artifacts go to `output/_quarantine/`. The authoritative path contract is `app/config/paths.py`.

Back up `templates/`, `template_store/`, `mappings/`, and `settings.json` when moving the application to another machine. Reports in the root `output/` folder can be regenerated from the original exports and staging workbooks.
