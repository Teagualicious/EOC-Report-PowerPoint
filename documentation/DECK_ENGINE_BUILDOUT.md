# Deck Engine Buildout Roadmap

This is the repository-facing execution plan distilled from the full Deck Engine fork handoff. `STATUS.md` records current state; this document defines the target architecture, ordered stages, gates, and non-negotiable controls.

## Product outcome

Deck Engine is a local-first Windows reporting tool with one analyst workflow:

1. Select one campaign export.
2. Parse it through a known import profile.
3. Generate an editable Excel staging workbook containing literal values.
4. Review and save that workbook in Excel.
5. Validate it.
6. Fill a pre-mapped PowerPoint template strictly from the saved workbook.

The donor is Jughead Data Engine v1.36.0. Preserve its parser, Unified Data, KPI, query, mapping, formatting, template-bundle, and PowerPoint-fill cores. Remove the MCP server, searchable/VBA workbook, COM live preview, old analyst windows, and other paths that compete with the six-step workflow.

## Architectural decisions

### The staging workbook is the source of truth

Parsing, alias resolution, KPI aggregation, queries, and campaign interpretation finish before the workbook is written. The deck builder reads literal values and image paths from the saved workbook; it must not consult live parsed data or require formula recalculation.

### Stable public workflow verbs

`app.engine.workflow` owns the application use cases:

- `parse_dump(path, profile=None)`
- `generate_staging(path, template=None, profile=None)`
- `validate_staging(path)`
- `build_deck(staging_path=None)`
- settings, state, and template-list operations

The CLI and UI remain thin shells. Engine modules never import Tk.

### Project-anchored paths and one sanitizer

All business paths derive from `app.config.paths`, never the process working directory. Every user-controlled filesystem name passes through `app.config.naming`; parallel sanitizers are prohibited.

### Synthetic data by default

No real client data, campaign exports, credentials, or internal reports may enter the repository. Tests use deterministic synthetic fixtures. First contact with a sanitized real export is a named acceptance risk, not permission to commit the file.

## Target module boundaries

- `parsers/`, `engine/data_pipeline.py`, `engine/excel_utils.py`: ingestion and Unified Data.
- `engine/kpi.py`, `metrics_catalog.py`, `query_resolver.py`, `pivot.py`: resolved reporting values.
- `engine/campaign_dictionary.py`: campaign normalization and interpretation.
- `engine/staging.py`: workbook writer/reader and contract versioning.
- `engine/validate.py`: pure validation rules and structured findings.
- `engine/pptx_formats.py`, `pptx_fill.py`, `fill_report.py`: deterministic deck output.
- `engine/pptx_mapper.py`, `template_bundle.py`, `mapper/`: developer-only template authoring.
- `engine/workflow.py`: sole public orchestration boundary.
- `ui/` and `cli.py`: presentation only.

## Staging workbook contract

The workbook will contain protected contract sheets and editable analyst surfaces. Exact sheet names may evolve during Stage 2, but the minimum contract is:

- metadata: contract version, source fingerprint, profile, template, timestamps, source hash;
- resolved values: key, literal value, type, format, provenance, validation state;
- campaign/placement detail sufficient for audit and approved edits;
- image references as normalized project-safe paths;
- validation findings and reconciliation totals.

Writer and reader must round-trip supported scalar types without formulas, macros, hidden live connections, or COM dependencies. Unknown contract versions fail loudly.

## Stage plan and gates

### Stage 0 — Fork surgery and architecture harness

Deliverables:

- establish the exact v1.36.0 baseline and survivor-test count;
- remove MCP, VBA/search export, COM preview, obsolete analyst windows, and dead dependencies;
- preserve parser/KPI/mapping/fill behavior;
- add stable workflow/module seams and project directories;
- enforce no-CWD paths, one sanitizer, and no-Tk-in-engine laws;
- leave the analyst UI intentionally unavailable.

Gate: survivor suite, architecture tests, compile check, and `python -m app.cli list-templates` pass.

### Stage 1 — Ingestion core

Deliverables:

- one-file CSV/XLSX/XLSM/HTML ingestion through `workflow.parse_dump`;
- stable structure fingerprint from normalized sheet names and header sets, independent of column order;
- stored `profile_<fingerprint>.json` import profiles;
- unknown fingerprints return a profile-required result and never guess;
- campaign-dictionary v0 identity passthrough with analyst notes;
- source reconciliation report and deterministic synthetic fixture factory;
- 50,000-row performance check.

Gate: known-profile replay is deterministic, column reordering retains the fingerprint, structural changes require a new profile, and malformed inputs produce actionable errors.

### Stage 2 — Staging workbook writer and reader

Deliverables:

- versioned workbook contract with literal values only;
- deterministic key ordering, provenance, reconciliation, and editable analyst fields;
- atomic writes and collision-safe filenames;
- reader that treats the saved workbook—not memory—as the next-stage input;
- round-trip tests for numbers, dates, text, blanks, booleans, images, and large tables.

Gate: write → close → reopen → read reproduces the contract exactly, with no formulas or external connections.

### Stage 3 — Validation gates

Deliverables:

- structured `error`, `warning`, and `info` findings;
- blocking checks for missing required sheets/keys, invalid contract versions, duplicate keys, unsafe paths, unresolved placeholders, non-finite numbers, malformed dates, and reconciliation breaks;
- warnings for stale sources, unusual variances, optional blanks, and unsupported formatting;
- quarantined output on failure.

Gate: every blocking rule has a focused test; warnings never silently become blockers; analyst messages identify location and remediation.

### Stage 4 — Fill integration

Deliverables:

- build only from a validated saved staging workbook;
- preserve template geometry, fonts, colors, number formats, persistent shape identity, chart formatting, and image placement;
- loud reporting of missing mappings or unmatched chart queries;
- atomic output and golden-deck semantic tests.

Gate: the builder cannot access parsed-data objects, the full inherited golden fill suite remains green, and a deterministic synthetic workbook produces the expected deck.

### Stage 5 — Analyst UI

Deliverables:

- one light-theme window with explicit states: no source, profile needed, ready to stage, staging generated, validation blocked/warned, ready to build, complete;
- source/template selection, profile editor, progress, findings, output links, and recovery actions;
- background work through a bounded worker bridge; no business logic in callbacks;
- keyboard, DPI, small-screen, cancellation, and file-lock handling.

Gate: an analyst completes the full workflow without terminal access; UI state is derived from workflow results rather than duplicate flags.

### Stage 6 — Settings, template management, and launcher

Deliverables:

- minimal settings for output, default template, and safe preferences;
- template import/list/remove with mapping-health checks;
- release-safe Windows launcher and portable package;
- development versions excluded from automatic releases.

Gate: clean-machine install and launcher checks pass on Windows; missing Office does not break non-rendering paths.

### Stage 7 — Documentation, UAT, parallel run, and rollout

Deliverables:

- analyst guide, maintenance guide, troubleshooting, data contract, release checklist, and recovery instructions;
- synthetic end-to-end UAT plus sanitized parallel comparison against the current process;
- recorded variances, approvals, rollback path, and ownership.

Gate: named analyst acceptance, no open blocking findings, and documented rollback.

### Stage 8 — Campaign dictionary v1 (optional)

Deliverables:

- explicit alias/rule model with precedence, confidence, and provenance;
- analyst-reviewed suggestions rather than silent inference;
- versioned migrations and regression fixtures.

Gate: ambiguous campaigns remain visible and unresolved unless an approved rule applies.

## Test strategy

Every behavior change receives a test. Required layers:

- unit tests for normalization, fingerprints, aliases, formatting, validation, and path safety;
- contract tests for workflow result shapes and staging versions;
- integration tests for parse → stage → reopen → validate → fill;
- semantic golden-deck tests using `python-pptx`, not binary file equality;
- architecture tests for dependency direction, path anchoring, sanitizer ownership, and source-of-truth isolation;
- performance checks for representative 50,000-row exports;
- manual Windows checks only where automation is not credible: Office rendering, DPI, launchers, locks, and clean-machine packaging.

Linux CI may not claim Windows/Office acceptance.

## Validation catalogue

Blocking examples:

- unreadable or unsupported source;
- unknown structure without an approved profile;
- missing contract sheet or required key;
- duplicate resolved key;
- unsupported contract version;
- formula or external link where literals are required;
- unsafe/escaping path;
- unresolved mapped placeholder;
- invalid numeric/date value;
- reconciliation outside approved tolerance;
- output collision that cannot be resolved atomically.

Warning examples:

- optional value missing;
- stale source timestamp;
- unusual KPI variance or empty campaign group;
- unused staging key or template mapping;
- optional image unavailable;
- analyst edit differs materially from source reconciliation.

## Risks and controls

- **RSK-1 — donor regression:** keep survivor/golden suites mandatory.
- **RSK-2 — synthetic fixtures miss a real export shape:** require sanitized first-contact acceptance before production rollout.
- **RSK-3 — workbook edits break reconciliation:** preserve provenance and validate before fill.
- **RSK-4 — template drift:** persistent shape identity and mapping-health checks.
- **RSK-5 — Windows-only failures:** explicit clean-machine and Office acceptance gates.
- **RSK-6 — silent inference:** unknown profiles and ambiguous campaign rules fail into analyst review.
- **RSK-7 — sensitive data leakage:** synthetic fixtures, ignore rules, and PR review.
- **RSK-8 — competing workflows return:** architecture tests and module-disposition review.

## Pull-request discipline

Each stage uses a dedicated branch and draft PR. A stage PR must include:

- scope and exclusions;
- changed contracts and migration notes;
- exact automated checks run and results;
- Windows/manual checks not performed;
- data-hygiene confirmation;
- `STATUS.md` update and the next stage’s entry condition.

Do not bump `VERSION` to a releasable value until the release checklist and Windows gate are complete.

## Definition of done

Deck Engine is complete when a clean Windows machine can ingest a supported export, request a profile instead of guessing, generate and reopen a literal staging workbook, surface blocking/warning findings, build a formatting-faithful mapped deck only from that workbook, retain auditable provenance/reconciliation, and recover from errors without exposing client data or requiring developer intervention.
