# Deck Engine Handoff

> **Authoritative onboarding and plan document.** State date: **2026-08-20**.
> Read together with [`../STATUS.md`](../STATUS.md) (current state, always read first) and
> [`DECK_ENGINE_BUILDOUT.md`](DECK_ENGINE_BUILDOUT.md) (stage gates and controls).
> Documentation rules: [`DOCUMENTATION_STANDARDS.md`](DOCUMENTATION_STANDARDS.md).
> Test rules: [`TESTING_STANDARDS.md`](TESTING_STANDARDS.md).

---

## 1. What this product is

Deck Engine is a **local-first Windows reporting tool** for Spectrum Reach end-of-campaign
(EOC) reporting. It is a **two-step program with a guiding dashboard**:

1. **Step 1 — dump → Excel.** Take one raw campaign export (CSV / XLSX / XLSM / HTML),
   combine and resolve its KPIs, and export them into an **editable Excel staging
   workbook** containing literal values.
2. **Step 2 — Excel → PowerPoint.** Fill a pre-mapped PowerPoint template **strictly from
   the saved staging workbook**, producing a formatting-faithful branded deck.

A **localhost web dashboard** (opened in the analyst's browser) guides the user through the
whole flow: pick a dump → generate the workbook → open it in desktop Excel to edit the
metrics → validate → build the deck. The staging workbook is the *only* editing surface for
report values and the *only* input to the deck build.

Until the real report template is supplied, a **default template using Spectrum Reach
branding** stands in; the code treats the template as replaceable configuration, not a
hardcoded artifact. From export to export the template changes only in small ways — e.g.
one page shows a **table** for one campaign and a **pie chart** for another. Per-page
content choices will be made with checkboxes on the dashboard (wireframed now, activated
when the real template exists — see DEC-3).

## 2. Lineage — how we got here

| When | What |
|---|---|
| ≤ 2026-07 | **Jughead Data Engine v1.36.0** (donor): a proven Tk desktop tool with parser, Unified Data, KPI, query, mapping, formatting, and PowerPoint-fill cores. 367 tests at the fork point. |
| 2026-08-10 | **Fork Stage 0** (done, PR #1 merged): stripped the MCP server, VBA/search workbook, COM live preview, and the multi-window analyst UI. Preserved the cores. Added stable `engine.workflow` verbs (`parse_dump`/`generate_staging`/`build_deck` stubs — note `validate_staging` was *not* stubbed; it is added in Stage 3), staging/validation/dictionary module contracts, path & sanitizer laws, and 4 architecture-enforcement tests. Survivor suite: **313 tests**. |
| 2026-08-20 | **This planning pass**: product decisions DEC-1…DEC-4 locked with the owner; target architecture and phase plan written down (this document); documentation and testing standards added. |

The donor documentation is archived under [`upstream/`](upstream/) and is **context, not
authority** — several subsystems it describes (MCP, VBA workbook, COM preview, old analyst
windows) no longer exist. `tests/test_architecture.py` fails CI if any stripped module
returns.

## 3. Product decisions (locked 2026-08-20 with the owner)

These four decisions came from the project owner and override anything older, including the
original Stage 5 text of `DECK_ENGINE_BUILDOUT.md` (now amended).

- **DEC-1 — The analyst UI is a localhost web dashboard.** A local web server bound to
  `127.0.0.1`, opened in the default browser, is the *only* analyst UI. The previously
  planned Tkinter one-window app is dropped. The engine stays headless; the dashboard is a
  thin shell over `engine.workflow`.
- **DEC-2 — Metrics are edited in real Excel.** The dashboard's "edit metrics" step opens
  the generated staging workbook in desktop Excel (`os.startfile`). The analyst edits and
  saves there; the dashboard then re-validates the saved file and unlocks "Build deck".
  There is **no** second (in-browser) editing surface — one source of truth, one editor.
- **DEC-3 — Per-page content checkboxes, wireframed now, live later.** Pages whose content
  varies (table vs pie chart, etc.) get per-page checkbox options on the dashboard so the
  analyst selects what appears. Until the real template exists we do not know which pages
  vary, so: build the UI wireframe and the plumbing, clearly marked, behind a feature flag
  that is **off by default**. Activate only when the actual template is available to
  reference.
- **DEC-4 — Synthetic-first continues.** No real client data, exports, or credentials in
  the repository, ever. Deterministic synthetic fixtures shaped like campaign exports drive
  all development. First contact with a sanitized real export is a named acceptance risk
  (RSK-2), not permission to commit the file.

## 4. Engineering decisions (this pass — revisable at the named gates)

- **DEC-5 — The template-first IR pipeline is the production deck-build carrier.**
  The repo contains **two complete mapping/fill systems** (see §6.3). Stage 4's
  `build_deck` will construct slot values from the staging workbook and build through
  `engine/template_ir/build.py` (`build_from_template`), because: it accepts pre-resolved
  values as its only per-build input (exactly the "workbook only" law), it re-emits
  verbatim shape XML, its reconciliation (Phase D) survives template revisions (risk
  RSK-4), and shape-level `excluded` is the natural extension point for DEC-3 page
  variants. **Known fidelity limit (RSK-12):** the builder re-emits shapes into a bare
  `Presentation()` — it does **not** carry the source theme, slide masters, layouts, or
  slide backgrounds. Fidelity is exact only for slide-level shapes with self-contained
  formatting (explicit `srgbClr` colors, explicit fonts); theme-referenced colors/fonts
  (`schemeClr`, `+mj-lt`/`+mn-lt`) silently resolve against the default Office theme, and
  master/layout-carried branding vanishes. Stage 4 must either (a) enforce
  self-contained authoring for templates (DEC-6 does this for the interim one) or
  (b) extend `build_from_template` to clone theme/master/background parts — decide at
  Stage 4 entry alongside the carrier confirmation, with a golden test using a
  `schemeClr` synthetic template either way. The classic mapper path stays green as the
  golden-regression harness and developer tool. **Revisit gate:** confirm at Stage 4
  entry; if reversed, record the reversal in `STATUS.md` decisions and update this
  section.
- **DEC-6 — Interim default template is a committed synthetic Spectrum-branded deck.**
  Author a generic branded template (cover, KPI summary, breakdown table page, breakdown
  chart page; navy palette from `app/config/themes.py`, logos from
  `app/resources/assets/`) with **no client data**, commit it plus its seeded template
  store and slot mapping. **Authoring constraints (per RSK-12):** all branding lives on
  the slides themselves — explicit RGB colors, explicit font names, no master/layout
  background art, no theme-color references. All code selects the template via
  settings/template-id, never by hardcoded filename, so swapping in the real template is:
  ingest → review/author slots → change the default. Gitignore work needed (two layers):
  `.gitignore` excludes the `workspace/template_store/` *directory* (files under an
  excluded directory cannot be re-included with a simple `!` on the file), and the global
  `*.xlsx` rule independently blocks the store's embedded chart workbooks
  (`assets/chart.xlsx`). The seeding commit must restructure the rules (e.g.
  `workspace/template_store/*` + `!…/default/` + `!…/default/**` +
  `!…/default/**/*.xlsx`) and add a check that the seeded files are actually tracked.
- **DEC-7 — Dashboard server is stdlib-first.** `http.server.ThreadingHTTPServer` +
  a small hand-rolled JSON router + static single-page frontend; no new runtime
  dependencies (repo law: standard library before new dependencies). Bind `127.0.0.1`
  only; require a random per-run token on every request (guards against drive-by/DNS-rebind
  requests to localhost); mutations are POST-only. Token transport: the bootstrap URL
  carries the token once; the frontend keeps it in memory/sessionStorage and sends it as a
  request header on every API call (URL/history exposure of the one-shot bootstrap token
  is accepted for a single-user localhost tool). If the stdlib server proves genuinely
  limiting, switching to Flask is allowed but must be recorded as a decision in
  `STATUS.md` first.
- **DEC-8 — Chart/table data becomes literal workbook content.** Pages filled as charts or
  tables need `{categories, series}` / `{headers, rows}` payloads. Because Stage 4 must
  not touch parsed data, these payloads are resolved during staging generation and written
  as **literal cells** in a dedicated `Chart & Table Data` sheet (it holds both payload
  kinds). This sheet is part of the **initial Stage 2 contract**: `SCHEMA_VERSION` is 1
  and no workbook has ever shipped, so no bump is needed for it — but any sheet or field
  added *after* Stage 2 closes (e.g. DEC-3 variant choices) requires a version bump and a
  loud unknown-version failure path. Editing the numbers there edits the chart.

## 5. Target user flow (what "done" behaves like)

```
Run Deck Engine.bat
  └─► starts dashboard server (127.0.0.1:<port>?token=…) ─► opens default browser

DASHBOARD (state machine, one state visible at a time)
  [no source] ──select dump (dashboard lists files dropped into input/)──► parse_dump
       │ unknown structure? ──► [profile needed] ─ profile editor ─┐
       ▼                                                           │
  [ready to stage] ──generate──► staging workbook in workspace/staging/
       ▼
  [staging generated] ──"Open in Excel"──► analyst edits metrics, saves, returns
       ▼            (future, flag off: per-page content checkboxes — DEC-3)
  [validate] ──findings──► [blocked] (errors listed w/ remediation) or [ready to build]
       ▼
  [ready to build] ──build──► deck in output/  ──► [complete: open deck / open folder]
                        └─ failure ──► output/_quarantine + loud findings
```

Everything the dashboard does goes through `engine.workflow` verbs; the CLI (`python -m
app.cli`) exposes the same verbs for scripting and tests. (The validate step's verb,
`workflow.validate_staging`, and its CLI command do **not** exist yet — they are a Stage 3
deliverable; Stage 0 stubbed only parse/stage/build.)

## 6. Architecture

### 6.1 Layering (enforced by `tests/test_architecture.py`)

```
┌────────────────────────── presentation (thin shells) ──────────────────────────┐
│ app/cli.py            JSON CLI (live)                                          │
│ app/dashboard/        localhost web dashboard (Stage 5, NEW — replaces Tk app) │
│ app/mapper/           developer-only Tk template mapper (live)                 │
│ app/ui/               residual Tk helpers used by the mapper                   │
└───────────────┬────────────────────────────────────────────────────────────────┘
                ▼  (presentation may import engine/config/parsers; never reverse)
┌────────────────────────── engine/ (headless, no Tk ever) ──────────────────────┐
│ workflow.py   THE ONLY public boundary: parse_dump · generate_staging ·        │
│               build_deck (stubs) · validate_staging (MISSING — added Stage 3) ·│
│               settings/state/templates · developer template-store verbs        │
│               (ingest_template_store · list_template_stores ·                  │
│                build_template_report — live, used to seed DEC-6)               │
│ data_pipeline · kpi · metrics_catalog · query_resolver · pivot · excel_utils   │
│ campaign_dictionary (S1) · staging (S2) · validate (S3)   ← contract stubs     │
│ pptx_mapper · pptx_formats · pptx_fill · fill_report · shape_identity ·        │
│ template_bundle · template_ir/ (ingest·classify·mapping·build·reconcile)       │
└───────────────┬────────────────────────────────────────────────────────────────┘
                ▼
┌── parsers/ ────────────────────────┐  ┌── config/ ─────────────────────────────┐
│ csv · excel · html · dictionary    │  │ paths (path law) · naming (sanitizer   │
│ (unified parsed-data dict)         │  │ law) · settings · themes · logging     │
└────────────────────────────────────┘  └────────────────────────────────────────┘
```

Laws (1–4 machine-enforced today by `tests/test_architecture.py`; law 5 from Stage 4):

1. **Path law** — every business path derives from `config.paths`; no `os.getcwd()`, no
   bare `Path()` (T-ARCH-1).
2. **Sanitizer law** — only `config/naming.py` defines filename sanitizers (T-ARCH-2
   enforces single *definition* ownership; actually routing every user-controlled name
   through it is review-enforced).
3. **Headless engine** — importing `engine.workflow/staging/validate/pptx_fill` loads zero
   tkinter (T-ARCH-3). Stage 5 adds: engine never imports `app.dashboard`; the dashboard
   never imports tkinter.
4. **Stripped modules stay dead** — the removed Jughead modules must not reappear
   (`test_stripped_runtime_modules_are_absent`; it carries no T-ARCH-4 docstring label yet
   — adding one is a recorded finding for the next code-touching pass).
5. **Source-of-truth law** — the saved staging workbook is the only input to
   `build_deck`; filling from live parsed data is prohibited (enforced from Stage 4).

### 6.2 Step 1 pipeline (dump → Excel) — data flow and current state

```
input/<dump>.(csv|xlsx|xlsm|html)
  → parsers.{CSVParser,ExcelParser,HTMLParser}          [LIVE donor core]
      one unified parsed-data dict per file: campaign_metrics, level_data,
      detected_tables, metrics…, columns classified via parsers/dictionary.py
      against app/resources/metric_dictionary.json (20 metric families,
      11 breakdown levels, sum-vs-avg aggregation law)
  → structure fingerprint → import profile profile_<fp>.json                [Stage 1]
      unknown fingerprint ⇒ "profile required", never guess
  → engine.campaign_dictionary.apply — v0 identity passthrough + notes      [Stage 1]
  → KPI resolution                                       [LIVE donor core]
      engine.kpi.compute_kpis (best-source-per-campaign law),
      engine.metrics_catalog.get_available_metrics (flat fill keys),
      engine.query_resolver / engine.pivot (queries, table/chart payloads)
  → engine.staging.write_staging_workbook                                   [Stage 2]
      workspace/staging/<name>.xlsx — literal values only
      sheets: Report Values · Images · Chart & Table Data (DEC-8) · Unified Data · _Meta
      (today staging.py has SCHEMA_VERSION=1 and four sheet constants; the payload
       sheet joins the contract when the Stage 2 writer ships — no bump needed, DEC-8)
  → engine.validate.validate_ingest/validate_fill → Finding list            [Stage 3]
```

KPI laws already encoded in the live cores and their tests (**do not "simplify"**):

- **Aggregation law**: `parsers.dictionary.get_metric_aggregation` is the single authority
  — `sum` by default, `avg` for ratio metrics (CTR, CPM, Frequency…). Registering a new
  rate metric there is mandatory *before* use, or additive summing silently corrupts it.
- **Best-source law**: additive metrics sum per (campaign, metric, source); the max source
  wins *per campaign*; campaign winners then sum. Per-campaign, not global-max.
- **Reach/Frequency** are omitted from review KPI totals (approved 2026-07-14);
  **Completion Rate** is derived (Σcompletions/Σvideo-starts), not averaged (2026-07-13).
- **Re-export replacement**: matching Unified-Data keys replace, others append
  (multi-period support).

### 6.3 Step 2 pipeline (Excel → PowerPoint) — two systems, one production path

Two complete mapping/fill systems exist; **do not conflate them**:

| | Classic mapper | Template-first IR (production path, DEC-5) |
|---|---|---|
| Storage | `workspace/mappings/pptx_<name>_pptx.json` | `workspace/template_store/<id>/{template.json, mapping.json, assets/, source.pptx}` |
| Model | per-slide → per-shape assignments (metric, format, replace_text) | shape XML stored **verbatim**; named slots (text/number/date/chart_data/table_data/image) in a slot registry |
| Fill | `engine/pptx_fill.py` edits the template copy in place | `engine/template_ir/build.py` re-emits verbatim XML into a fresh deck, then fills slots |
| Fidelity | preserves runs it touches | pixel-exact by construction (never re-authors shapes) |
| Template revisions | manual re-mapping | `reconcile.py` carries review work across re-ingest, loud deltas |
| Role going forward | golden-regression harness + developer Quick Fill | **Stage 4 `build_deck` carrier** |

Shared machinery: shape identity (`shape_identity.py`: uid → unique name → positional only
for legacy; ambiguity = skip loudly), formatting (`pptx_formats.py`, the single formatting
authority incl. numpy coercion), text replacement, chart injection (whole chart part cloned
— XML + embedded workbook + colors — then `replace_data`), and the loud
`FillReport`/build-report convention: **a fill never silently partially succeeds**.

Stage 4 wiring (the missing piece): `workflow.build_deck(staging_path)` =
`read_staging_workbook` → assemble `slot_values` → `build_from_template` → atomic write to
`output/`, quarantine + findings on failure. No parsed-data objects anywhere in that call
path. Three contract details a literal implementation must not skip:

- **Slot keys.** `mapping.json` slots are *queries* that today re-resolve over parsed
  `client_data` — forbidden in Stage 4. Therefore `generate_staging` resolves **every**
  slot of the active template's mapping via `query_resolver`/`pivot` at staging time and
  writes one workbook row per slot, **keyed by slot name** (scalars → `Report Values`,
  payloads → `Chart & Table Data`, image refs → `Images`), recording the mapping hash in
  `_Meta`. `build_deck` then maps rows to slots by key with zero query resolution.
- **Formatting.** `build_from_template`'s text fill is `str(value)`; in the existing path
  `resolve_slot_values` formats display strings via `engine.pptx_formats` *before* the
  build. So either the workbook stores display-ready strings, or Stage 4 applies each
  row's format/format_details columns through `pptx_formats` during assembly — Q1 decides
  which at Stage 2 entry. Raw scalars fed straight in would ship `1234567.0` decks.
- **Images.** `_apply_image` takes a `{image_path, image_path_abs}` payload dict, not a
  bare path — `Images` rows are wrapped accordingly, with absolute paths resolved against
  `config.paths`.

### 6.4 The dashboard (Stage 5, new `app/dashboard/` package)

- `server.py` — stdlib HTTP plumbing (DEC-7): 127.0.0.1 bind, ephemeral port, per-run
  token check, static file serving, JSON dispatch. No business logic.
- `api.py` — endpoint handlers, each a thin call into `engine.workflow` returning its
  JSON-native result (same thin-shell rule as the CLI, enforced by test).
- `static/` — single-page frontend (`index.html`, `app.js`, `styles.css`); state machine
  rendering of `describe_state` + step results; polling, no websockets.
- **Flow-state ownership:** today's `describe_state()` returns only
  version/phase/paths/settings/templates — it knows nothing of a selected dump, a
  generated workbook, or validation status, and Phase 5 forbids duplicate state flags in
  the frontend. Stage 5 therefore **extends `describe_state()` to derive flow state from
  durable artifacts** (latest staging workbook + its persisted validation findings), keeping
  the engine stateless and the workbook the only carrier; the extended shape is a
  documented contract. Server-side memory of parsed results is prohibited — it is exactly
  the parsed-data shortcut the source-of-truth law bans.
- `__main__.py` — `python -m app.dashboard`; `app/main.py` becomes: ensure dirs → start
  server → `webbrowser.open` (so `python -m app` and the .bat launcher are the analyst
  entry points from Stage 5 on).
- Variant panel (DEC-3): renders per-page checkbox groups from a `variant_manifest`
  (empty until the real template); choices are POSTed and stored **in the staging
  workbook** so builds remain workbook-driven; behind `page_variants` feature flag,
  default off, wireframe visibly labeled as future. Windows file-lock reality: while
  Excel holds the workbook open (the DEC-2 step), an openpyxl save will raise
  `PermissionError` — so variant choices are held server-side and written only when the
  file is unlocked, with a visible retry message on failure.
- Open-in-Excel (DEC-2): `os.startfile(staging_path)` on Windows; non-Windows dev fallback
  logs the path. Re-validation triggered by the analyst ("I've saved — validate") and by
  mtime polling as a convenience; because Excel saves via temp-file-then-rename, an
  mtime-triggered read can catch a transient state — validation retries on open/parse
  failure before reporting anything.

### 6.5 Directory contract (from `config.paths` — the code, not the stale workspace README)

```
project root
├─ input/                     analyst drops raw dumps here (gitignored)
├─ output/                    finished decks; output/_quarantine for rejected artifacts
├─ workspace/                 application-managed state
│  ├─ staging/                staging workbooks (Stage 2)   [gitignored]
│  ├─ templates/ (+images/)   saved templates (StandIn_Report.pptx committed)
│  ├─ template_store/<id>/    template-first stores          [gitignored; DEC-6 seeds one]
│  │                          (path defined as TEMPLATE_STORE_DIR in
│  │                           engine/template_ir/ingest.py, derived from
│  │                           config.paths.WORKSPACE_DIR — not in config.paths itself,
│  │                           and not created by ensure_dirs())
│  ├─ mappings/               classic mapping JSON + profile_/platform_/fingerprint_ JSON
│  ├─ dictionary/             generated dictionary working files (reserved)
│  ├─ logs/                   rotating logs + fill_history.jsonl
│  └─ settings.json           atomic-write settings
└─ app/resources/             read-only: metric_dictionary.json, brand assets
```

Gotcha: `.gitignore` ignores `workspace/mappings/` yet two mapping JSONs are
**tracked** (committed before the rule; the index wins) — they act as seeds; new mapping
files stay local unless force-added deliberately.

## 7. Phase plan

Phases keep the buildout's stage numbers (the codebase, STATUS.md, and NotImplementedError
messages all reference them). Every phase closes per the session workflow: tests green →
commit → `STATUS.md` updated with what changed / checks run / decisions / next entry
condition → branch + draft PR. **Gates are owned by
[`DECK_ENGINE_BUILDOUT.md`](DECK_ENGINE_BUILDOUT.md)** — they are not restated here; each
phase below lists its objective and only the *deltas* this plan adds on top of that
stage's spec.

### Phase 0 — Fork surgery & architecture harness ✅ DONE (PR #1)
313-test survivor suite green; laws enforced; parse/stage/build verbs stubbed.

### Phase 1 — Ingestion core (Step 1, part A) — buildout Stage 1
**Objective:** one dump in, deterministic structured data out — never guessing.
Deltas: the synthetic fixture factory is the shared foundation every later phase extends
(TESTING_STANDARDS §3); the import-profile *schema* gets defined and documented here
(feeds Q6 — the Stage 5 profile editor edits this contract).

### Phase 2 — Staging workbook writer/reader (Step 1, part B) — buildout Stage 2
**Objective:** the literal, editable, versioned workbook that is the product's center —
this completes the owner's "dump → combined KPIs → Excel" step.
Deltas:
- The `Chart & Table Data` sheet ships in the initial contract (DEC-8).
- **Slot resolution moves to staging time:** `generate_staging` resolves every slot of the
  active template's `mapping.json` via `query_resolver`/`pivot` and writes one row per
  slot, keyed by slot name, with the mapping hash in `_Meta` (see §6.3) — this is what
  makes a query-free Stage 4 possible.
- Q1 (raw value + format columns vs display-ready strings) is decided at entry.

### Phase 3 — Validation gates — buildout Stage 3
**Objective:** nothing broken reaches PowerPoint; every failure tells the analyst where
and how to fix it.
Deltas:
- **Add `workflow.validate_staging(path)`** delegating to `engine.validate`, plus a CLI
  `validate` subcommand, with contract tests — the verb does not exist today (not even a
  stub) and the §5 flow and Phase 5 dashboard depend on it.
- Quarantine flow into `output/_quarantine`.

### Phase 4 — Deck build from workbook (Step 2) + interim branded template — buildout Stage 4
**Objective:** `build_deck(staging_path)` produces the deck from the workbook alone.
Deltas:
- Confirm DEC-5 at entry, including the RSK-12 theme/master decision (author-constrained
  templates vs cloning theme parts) with a `schemeClr` golden test either way.
- Wiring per §6.3: slot-keyed rows, formatting through `pptx_formats`, image payload
  wrapping; loud report of unmapped/unfilled slots.
- Author + commit the Spectrum-branded default template, seeded store (gitignore
  restructuring per DEC-6), slot mapping; document the replace-the-template procedure.
- Architecture test extension: `build_deck` call path provably free of parsed-data objects.

### Phase 5 — Localhost dashboard (replaces the Tk analyst UI — DEC-1) — buildout Stage 5
**Objective:** an analyst completes the entire flow without a terminal.
Deltas:
- `app/dashboard/` per §6.4, including the **flow-state extension of `describe_state()`**
  (derived from durable artifacts; documented contract) and the file-lock/mtime-retry
  behaviors.
- Profile editor scope per Q6; variant checkbox wireframe with flag off (DEC-3);
  Open-in-Excel (DEC-2).
- Fix the `python -m app` bootstrap bug (recorded in STATUS.md) as part of making
  `app/main.py` the dashboard entry point.
- Endpoint logic tests run headless against the handler layer; server-behavior tests
  (bind/token/POST-only) use loopback sockets per TESTING_STANDARDS §4.

### Phase 6 — Settings, template management, launcher — buildout Stage 6
Deltas: settings surface lives in the dashboard; template health checks cover both mapping
systems; the launcher starts the server and opens the browser. Open questions Q3–Q5
(dead COM remnants, template-store bundling, tracked-mapping seeds) are decided here.

### Phase 7 — Documentation, UAT, parallel run, rollout — buildout Stage 7
Deltas: none — as specified, including retiring RSK-2 via the sanitized parallel
comparison and recording the executed Windows acceptance checklist.

### Phase 8 — Campaign dictionary v1 (optional) — buildout Stage 8

### Future feature (explicitly parked): live page variants
Entry condition: **the real report template is in hand.** Then: identify variant pages;
design the variant model on template-first `excluded`-at-build-time (or slide-level
variants if the template demands it — new `SlideIR` field + builder/validator/reconciler
awareness); make `validate_slot_mapping` variant-aware (inactive variant's slots must not
warn); persist variant choices as staging fields — a post-Stage-2 schema change, so it
follows DEC-8's `SCHEMA_VERSION` bump rule; flip the `page_variants` flag; add golden
variant-build tests. Until then the dashboard shows the wireframe only.

## 8. Risks (carried + new)

RSK-1…RSK-8 from the buildout still stand. Added this pass:

- **RSK-9 — localhost attack surface.** A local HTTP server is reachable by any local
  process and, without the token check, by hostile web pages. Controls: DEC-7 (bind +
  token + POST-only), sanitizer law on all inputs, architecture test coverage.
- **RSK-10 — dual mapping systems drift.** Two fill systems double the maintenance
  surface until DEC-5 is confirmed and the classic path is formally scoped to
  regression/dev duty. Control: decision gate at Phase 4 entry; shared-machinery tests.
- **RSK-11 — Excel round-trip edge cases.** Analysts will do unexpected things in Excel
  (formulas, formats, moved rows), and Windows Excel holds a write lock on the open
  workbook while saving via temp-file-then-rename. Controls: Stage 3 blocking rules
  (formula/external-link detection, key checks), provenance columns, quarantine, and the
  §6.4 lock/retry behaviors (server-side hold for writes while locked; validation retries
  on transient open/parse failure).
- **RSK-12 — template-first theme/master fidelity gap.** `build_from_template` re-emits
  shapes into a bare `Presentation()` with no source theme, masters, layouts, or slide
  backgrounds — theme-referenced formatting silently falls back to the default Office
  theme. Controls: DEC-6 authoring constraints for the interim template; explicit Stage 4
  entry decision (constrain authoring vs clone theme parts); a `schemeClr` golden test
  that fails loudly if the gap ever bites.

## 9. How to pick this up (first 30 minutes of any session)

1. Read `STATUS.md` → this file → `DECK_ENGINE_BUILDOUT.md` for your stage's gate.
   `CLAUDE.md` is session law; the standards docs govern how you document and test.
2. Verify the baseline before changing anything:
   ```bash
   python -m pip install -r requirements.txt
   python -m compileall -q app tests
   python -m pytest -q        # expect the count pinned in STATUS.md
   python -m app.cli state    # JSON snapshot of paths/version/readiness
   ```
   On Linux without tkinter, `tests/test_ui_helpers.py` and
   `tests/test_query_builder_helpers.py` fail on import — an environment gap, not a
   regression; CI has tkinter.
3. Work **only** the stage named in `STATUS.md` → *Next*. Unrelated findings go into
   `STATUS.md`, not into the diff.
4. What runs today: `python -m app.cli list-templates|state|settings`, plus the hidden
   developer commands `python -m app.cli ingest-template --pptx …` and `template-stores`
   (suppressed from `--help`; they drive the template-store tooling DEC-6 seeding uses).
   The developer mapper runs via `python -m app.mapper` (Windows). `parse`/`stage`/`build`
   raise `NotImplementedError` until their stages land; there is no `validate` command yet
   (Stage 3). **Known bug:** `python -m app` does *not* exit 2 as `app/main.py` intends —
   it crashes with `ModuleNotFoundError: No module named 'config'` (exit 1) because
   `app/__main__.py` lacks the `sys.path` bootstrap `app/cli.py` has; recorded in
   STATUS.md, fix scheduled with the Phase 5 entry-point rework.
5. Close the session per the workflow in `CLAUDE.md` (tests → commit → STATUS.md →
   branch + draft PR).

## 10. Open questions (tracked; answer before or at the named stage)

| # | Question | Decide at |
|---|---|---|
| Q1 | The `Report Values` key contract: donor catalog names (`Total Impressions`, `Avg CTR`, `zip:90210:Impressions`) as-is or a versioned namespace (donor names recommended — mapping files already use them) — **and** whether rows store raw value + format/format_details columns or display-ready strings (see §6.3 formatting note). | Stage 2 entry |
| Q2 | Exact real-dump shape (sheets/columns/KPI set) — synthetic fixtures approximate it until a sanitized export arrives (RSK-2). | Stage 1 → Stage 7 |
| Q3 | Should dead COM/live-preview remnants in `app/mapper/` + `pptx_thumbs.py` be removed or kept for a possible Windows re-enable? (Currently harmless, guarded.) | Stage 5/6 |
| Q4 | Template-store portability: extend `template_bundle` to bundle template stores, or keep bundles classic-only? | Stage 6 |
| Q5 | The two tracked mapping JSONs vs the `workspace/mappings/` ignore rule — formalize as committed seeds or untrack? | Stage 6 |
| Q6 | Import-profile contract and editor scope: today `save_import_profile` stores only `{schema_version, fingerprint}`; Stage 1 defines the real schema (column roles? sheet selection?), and the Stage 5 profile editor's scope follows from it. | Schema: Stage 1 · editor: Stage 5 entry |
