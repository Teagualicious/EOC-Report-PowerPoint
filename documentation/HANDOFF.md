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
| 2026-08-10 | **Fork Stage 0** (done, PR #1 merged): stripped the MCP server, VBA/search workbook, COM live preview, and the multi-window analyst UI. Preserved the cores. Added stable `engine.workflow` verbs (stubs for the pipeline), staging/validation/dictionary module contracts, path & sanitizer laws, and 4 architecture-enforcement tests. Survivor suite: **313 tests**. |
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
  values as its only per-build input (exactly the "workbook only" law), it preserves
  fidelity by re-emitting verbatim shape XML, its reconciliation (Phase D) survives
  template revisions (risk RSK-4), and shape-level `excluded` is the natural extension
  point for DEC-3 page variants. The classic mapper path stays green as the
  golden-regression harness and developer tool. **Revisit gate:** confirm at Stage 4 entry;
  if reversed, record the reversal in `STATUS.md` decisions and update this section.
- **DEC-6 — Interim default template is a committed synthetic Spectrum-branded deck.**
  Author a generic branded template (cover, KPI summary, breakdown table page, breakdown
  chart page; navy palette from `app/config/themes.py`, logos from
  `app/resources/assets/`) with **no client data**, commit it plus its seeded template
  store and slot mapping. All code selects the template via settings/template-id, never by
  hardcoded filename, so swapping in the real template is: ingest → review/author slots →
  change the default. (Note: `workspace/template_store/` is currently gitignored; the
  seeded default store needs an explicit negation rule.)
- **DEC-7 — Dashboard server is stdlib-first.** `http.server.ThreadingHTTPServer` +
  a small hand-rolled JSON router + static single-page frontend; no new runtime
  dependencies (repo law: standard library before new dependencies). Bind `127.0.0.1`
  only; require a random per-run token on every request (guards against drive-by/DNS-rebind
  requests to localhost); mutations are POST-only. If the stdlib server proves genuinely
  limiting, switching to Flask is allowed but must be recorded as a decision in
  `STATUS.md` first.
- **DEC-8 — Chart/table data becomes literal workbook content.** Pages filled as charts or
  tables need `{categories, series}` / `{headers, rows}` payloads. Because Stage 4 must
  not touch parsed data, these payloads are resolved during staging generation and written
  as **literal cells** in a dedicated `Chart Data` sheet of the staging workbook (schema
  addition; bump `engine/staging.py` `SCHEMA_VERSION` when it lands — target Stage 2, at
  the latest Stage 4 entry). Editing the numbers there edits the chart.

## 5. Target user flow (what "done" behaves like)

```
Run Deck Engine.bat
  └─► starts dashboard server (127.0.0.1:<port>?token=…) ─► opens default browser

DASHBOARD (state machine, one state visible at a time)
  [no source] ──select dump (from input/ or upload)──► parse_dump
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
app.cli`) exposes the same verbs for scripting and tests.

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
│               validate_staging · build_deck · settings/state/templates         │
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

Laws (all machine-enforced today, extended in Stage 5):

1. **Path law** — every business path derives from `config.paths`; no `os.getcwd()`, no
   bare `Path()` (T-ARCH-1).
2. **Sanitizer law** — only `config/naming.py` defines filename sanitizers; every
   user-controlled filesystem name goes through it (T-ARCH-2).
3. **Headless engine** — importing `engine.workflow/staging/validate/pptx_fill` loads zero
   tkinter (T-ARCH-3). Stage 5 adds: engine never imports `app.dashboard`; the dashboard
   never imports tkinter.
4. **Stripped modules stay dead** — the removed Jughead modules must not reappear
   (T-ARCH-4 in `test_architecture.py`).
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
      workspace/staging/<name>.xlsx — literal values only, SCHEMA_VERSION 1
      sheets: Report Values · Images · Chart Data (DEC-8) · Unified Data · _Meta
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
`read_staging_workbook` → assemble `slot_values` (scalars from `Report Values`, payloads
from `Chart Data`, paths from `Images`) → `build_from_template` → atomic write to
`output/`, quarantine + findings on failure. No parsed-data objects anywhere in that call
path.

### 6.4 The dashboard (Stage 5, new `app/dashboard/` package)

- `server.py` — stdlib HTTP plumbing (DEC-7): 127.0.0.1 bind, ephemeral port, per-run
  token check, static file serving, JSON dispatch. No business logic.
- `api.py` — endpoint handlers, each a thin call into `engine.workflow` returning its
  JSON-native result (same thin-shell rule as the CLI, enforced by test).
- `static/` — single-page frontend (`index.html`, `app.js`, `styles.css`); state machine
  rendering of `describe_state` + step results; polling, no websockets.
- `__main__.py` — `python -m app.dashboard`; `app/main.py` becomes: ensure dirs → start
  server → `webbrowser.open` (so `python -m app` and the .bat launcher are the analyst
  entry points from Stage 5 on).
- Variant panel (DEC-3): renders per-page checkbox groups from a `variant_manifest`
  (empty until the real template); choices are POSTed and stored **in the staging
  workbook** so builds remain workbook-driven; behind `page_variants` feature flag,
  default off, wireframe visibly labeled as future.
- Open-in-Excel (DEC-2): `os.startfile(staging_path)` on Windows; non-Windows dev fallback
  logs the path. Re-validation triggered by the analyst ("I've saved — validate") and by
  mtime polling as a convenience.

### 6.5 Directory contract (from `config.paths` — the code, not the stale workspace README)

```
project root
├─ input/                     analyst drops raw dumps here (gitignored)
├─ output/                    finished decks; output/_quarantine for rejected artifacts
├─ workspace/                 application-managed state
│  ├─ staging/                staging workbooks (Stage 2)   [gitignored]
│  ├─ templates/ (+images/)   saved templates (StandIn_Report.pptx committed)
│  ├─ template_store/<id>/    template-first stores          [gitignored; DEC-6 seeds one]
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
condition → branch + draft PR. Gates below are *additive* to
[`DECK_ENGINE_BUILDOUT.md`](DECK_ENGINE_BUILDOUT.md); that file remains the gate authority.

### Phase 0 — Fork surgery & architecture harness ✅ DONE (PR #1)
313-test survivor suite green; laws enforced; pipeline verbs stubbed.

### Phase 1 — Ingestion core (Step 1, part A)
**Objective:** one dump in, deterministic structured data out — never guessing.
- `workflow.parse_dump` over the live parsers; structure fingerprint from normalized sheet
  names + header sets (column-order independent); `profile_<fingerprint>.json` import
  profiles; unknown fingerprint ⇒ `profile_required` result.
- `campaign_dictionary.apply` v0 identity passthrough with analyst notes.
- Deterministic synthetic fixture factory (the shapes CI builds on for every later phase);
  source reconciliation report; 50,000-row performance check.
- **Gate:** known-profile replay deterministic; reorder-columns keeps fingerprint;
  structural change demands a new profile; malformed input → actionable `ParserError`.

### Phase 2 — Staging workbook writer/reader (Step 1, part B — completes "dump → Excel")
**Objective:** the literal, editable, versioned workbook that is the product's center.
- Implement `staging.write_staging_workbook` / `read_staging_workbook` /
  `inventory_mapping` against SCHEMA_VERSION (bump when `Chart Data` lands, DEC-8).
- Deterministic key ordering; provenance + reconciliation columns; editable analyst
  fields; atomic writes; collision-safe names via `config.naming`.
- **Gate:** write → close → reopen → read round-trips every supported type (numbers,
  dates, text, blanks, booleans, images, large tables, chart payloads) with **no formulas,
  macros, or external connections**; unknown contract version fails loudly.

### Phase 3 — Validation gates
**Objective:** nothing broken reaches PowerPoint; every failure tells the analyst where and
how to fix it.
- Implement `validate.validate_ingest` / `validate_fill` returning structured `Finding`s
  (severity/code/sheet/row/column/message/remedy) per the buildout's validation catalogue.
- Quarantine flow into `output/_quarantine`.
- **Gate:** every blocking rule has a focused test; warnings never silently escalate.

### Phase 4 — Deck build from workbook (Step 2) + interim branded template
**Objective:** `build_deck(staging_path)` produces the deck from the workbook alone.
- Confirm DEC-5 at entry. Wire workbook → `slot_values` → `build_from_template`; atomic
  output; loud report of unmapped/unfilled slots.
- Author + commit the Spectrum-branded default template, seeded store, slot mapping
  (DEC-6); document the replace-the-template procedure.
- Architecture test extension: `build_deck` call path provably free of parsed-data objects.
- **Gate:** inherited golden fill suite still green; deterministic synthetic workbook →
  expected deck; builder cannot see parsed data.

### Phase 5 — Localhost dashboard (replaces the Tk analyst UI — DEC-1)
**Objective:** an analyst completes the entire flow without a terminal.
- `app/dashboard/` per §6.4; explicit UI states mirroring workflow results (no duplicate
  state flags in the frontend); profile editor; findings panel with remediation text;
  recovery actions; Open-in-Excel (DEC-2); variant checkbox wireframe, flag off (DEC-3).
- Security: localhost bind + token + POST-only mutations + sanitizer on every
  user-supplied name (tested).
- New architecture tests: engine ⊬ dashboard; dashboard ⊬ tkinter; endpoints thin.
- Endpoint tests run headless against the handler layer (no browser, no port).
- **Gate:** full synthetic flow driven through HTTP endpoints in tests; UI state derived
  solely from workflow results.

### Phase 6 — Settings, template management, launcher
- Minimal settings surface in the dashboard; template import/list/remove with
  mapping-health checks (both systems); `Run Deck Engine.bat` → start server + open
  browser; portable package updated; dev VERSIONs still never release.
- **Gate:** clean-machine Windows install: launcher opens the dashboard; missing Office
  breaks nothing except actual Office rendering.

### Phase 7 — Documentation, UAT, parallel run, rollout
- Analyst guide, maintenance guide, troubleshooting, data contract, release checklist,
  recovery instructions; synthetic end-to-end UAT; sanitized parallel comparison against
  the current manual process (RSK-2 retired here); Windows acceptance checklist executed
  and recorded.
- **Gate:** named analyst acceptance; no open blocking findings; documented rollback.

### Phase 8 — Campaign dictionary v1 (optional, as specified in the buildout)

### Future feature (explicitly parked): live page variants
Entry condition: **the real report template is in hand.** Then: identify variant pages;
design the variant model on template-first `excluded`-at-build-time (or slide-level
variants if the template demands it — new `SlideIR` field + builder/validator/reconciler
awareness); make `validate_slot_mapping` variant-aware (inactive variant's slots must not
warn); flip the `page_variants` flag; add golden variant-build tests. Until then the
dashboard shows the wireframe only.

## 8. Risks (carried + new)

RSK-1…RSK-8 from the buildout still stand. Added this pass:

- **RSK-9 — localhost attack surface.** A local HTTP server is reachable by any local
  process and, without the token check, by hostile web pages. Controls: DEC-7 (bind +
  token + POST-only), sanitizer law on all inputs, architecture test coverage.
- **RSK-10 — dual mapping systems drift.** Two fill systems double the maintenance
  surface until DEC-5 is confirmed and the classic path is formally scoped to
  regression/dev duty. Control: decision gate at Phase 4 entry; shared-machinery tests.
- **RSK-11 — Excel round-trip edge cases.** Analysts will do unexpected things in Excel
  (formulas, formats, moved rows). Controls: Stage 3 blocking rules (formula/external-link
  detection, key checks), provenance columns, quarantine.

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
4. What runs today: `python -m app.cli list-templates|state|settings`; developer mapper via
   `python -m app.mapper` (Windows). `parse`/`stage`/`build` raise `NotImplementedError`
   until their stages land. `python -m app` exits 2 by design until Phase 5.
5. Close the session per the workflow in `CLAUDE.md` (tests → commit → STATUS.md →
   branch + draft PR).

## 10. Open questions (tracked; answer before or at the named stage)

| # | Question | Decide at |
|---|---|---|
| Q1 | Which flat-metric keys become `Report Values` keys — the donor catalog names (`Total Impressions`, `Avg CTR`, `zip:90210:Impressions`) as-is, or a versioned key namespace? Donor names recommended (mapping files already use them). | Stage 2 entry |
| Q2 | Exact real-dump shape (sheets/columns/KPI set) — synthetic fixtures approximate it until a sanitized export arrives (RSK-2). | Stage 1 → Stage 7 |
| Q3 | Should dead COM/live-preview remnants in `app/mapper/` + `pptx_thumbs.py` be removed or kept for a possible Windows re-enable? (Currently harmless, guarded.) | Stage 5/6 |
| Q4 | Template-store portability: extend `template_bundle` to bundle template stores, or keep bundles classic-only? | Stage 6 |
| Q5 | The two tracked mapping JSONs vs the `workspace/mappings/` ignore rule — formalize as committed seeds or untrack? | Stage 6 |
