# Deck Engine AI context

The authoritative plan is `documentation/DECK_ENGINE_BUILDOUT.md`; current state is `STATUS.md`.

Core invariant: **the saved staging workbook is the only source of truth for the PowerPoint fill.** Parsing, aliases, KPI aggregation, queries, and campaign rules resolve when staging is generated. Building a deck performs deterministic key→literal substitution plus image placement and must not consult live parsed data.

Inherited load-bearing modules:

- `app/parsers/*`, `engine/data_pipeline.py`, `engine/excel_utils.py`
- `engine/kpi.py`, `metrics_catalog.py`, `query_resolver.py`, `pivot.py`
- `engine/pptx_formats.py`, `pptx_fill.py`, `fill_report.py`, `shape_identity.py`
- `engine/pptx_mapper.py`, `template_bundle.py`; `app/mapper/*` is developer-only
- `config.paths`, `settings`, `naming`, `logging_setup`

New seams:

- `engine/campaign_dictionary.py` — campaign interpretation, identity passthrough in Stage 1
- `engine/staging.py` — workbook contract, Stage 2
- `engine/validate.py` — pure rules, Stage 3
- `engine/workflow.py` — only public verbs: parse, stage, build, state/settings/templates

Removed runtime surfaces: MCP, searchable Excel/VBA, PowerPoint live COM preview, old main/review/client/settings windows, dark theme, pywin32, tkcalendar.
