# Review: Template-First Mapper Architecture (proposal of 2026-07-15)

Reviews `../proposals/TEMPLATE_FIRST_MAPPER_2026-07-15.md` — the owner's
proposed Ingest → Classify → Map → Build pipeline that replaces in-place
template editing with a JSON intermediate representation (IR) and a
from-scratch deck build. **Status: accepted direction, not scheduled.**
A future session starting this work should read the proposal, then this
review, then the "How it maps onto the existing codebase" table below.

## Verdict

The architecture is sound and worth building — with one structural
amendment (static shapes must be copied as verbatim XML, not rebuilt from
the schema; see Critique 1). The core insight is correct: moving the
messy work (parsing arbitrary client decks) to a one-time, human-reviewed
ingest step, and making the monthly operation a clean build from a known
schema, converts recurring runtime fragility into one-time setup effort.
It also completes two threads this project already started:

- **Stable identity** — named slots finish what mapper roadmap Phase 4
  (`shape_uid`) began: durable, human-meaningful references instead of
  positions.
- **AI-nativeness** — a JSON IR + slot registry is exactly the shape of
  artifact Claude can read, validate, diff, and author mappings against
  via the existing MCP/CLI layer. This is the strongest synergy in the
  proposal and deserves to stay a first-class requirement.

## What the proposal gets right (keep as-is)

- EMUs everywhere, no unit conversion in storage.
- Run-level text extraction (the two-runs-one-box KPI callout analysis is
  accurate — that is precisely how these decks are built).
- Slot registry with types/descriptions → map-time validation ("text
  column mapped to a number slot") catches at map time what today
  surfaces as a wrong-looking shipped deck.
- Assets on disk, referenced by path — not base64 in JSON.
- Auto-classification as *suggestions* with a mandatory human review GUI.
- Honest edge-case list (SmartArt, theme colors, master inheritance,
  autofit) — these are the real dragons.
- Implementation order that defines the schema contract first.

## Critiques and required design changes

### 1. Static shapes: copy the original XML verbatim — never rebuild them
The proposal's biggest fidelity risk is rebuilding *branding* from the
schema. The schema inevitably captures a subset of OOXML: shadows,
glow/reflection effects, gradient details, freeform/custom geometry,
picture borders and rounded corners, transparency, text-frame insets,
word-wrap/column settings, kerning, bullet formats, hyperlinks… every
property the schema misses is visible brand drift on every generated
deck, discovered one client at a time (the proposal's Step 6 "fidelity
testing" is the tell — that gap-hunt never converges).

The proposal's own Option B contains the fix; make it central instead of
a fallback: **at ingest time, store each shape's original `<p:sp>` XML
verbatim alongside its parsed schema entry** (`element_xml` field or a
sidecar file). At build time:

- **static shapes** → deep-copy the stored XML into the new slide's
  `spTree`. Pixel-perfect branding, guaranteed, zero schema coverage
  needed. (lxml deepcopy into the shape tree is a known-reliable
  python-pptx technique; images additionally need their relationship
  re-targeted to the copied asset part.)
- **dynamic shapes** → build from the schema (they change every month
  anyway; this is where schema-driven construction earns its keep), OR
  copy-then-substitute like Option B for text-only slots.

This inverts the risk profile: fidelity risk exists only where data
changes — exactly where a human looks every month. The parsed schema for
static shapes remains valuable as *preview/documentation/diff* data, but
is never the render source.

### 2. Charts are the hardest part — plan them as their own phase
python-pptx chart *creation* covers the basic types but loses styling
nuance (data-label placement, gap width/overlap, axis number formats,
rounded corners on bars). The reliable path mirrors Critique 1: clone the
original chart part (its XML plus the embedded workbook part and
relationships — more involved than shape XML) and then update its cached
data + workbook, which is the same operation the current COM path already
does live. Budget charts as a dedicated phase with their own golden
tests; do not gate the first release on them.

### 3. Slots must survive re-ingest (template evolution)
`run_index: 0` and ids like `"slide3_shape7"` are positional — the exact
fragility Phase 4 just eliminated. Requirements:

- Schema shape ids must incorporate the persistent PowerPoint
  `shape.shape_id` (the Phase 4 `shape_uid`), with shape name as
  secondary identity.
- **Re-ingest reconciliation is a feature, not an afterthought**: when a
  client updates their template, re-ingest and carry classifications and
  slot names forward by uid → name matching; present only the *deltas*
  for review. Without this, every template revision costs a full
  re-classification session, and adoption dies there.
- Compound slots should pin to (shape uid, run **role**) with the run's
  placeholder text as the reconciliation hint, not a bare run index.

### 4. Theme colors: store the reference AND the resolved hex
Resolving "Accent 1" to hex at ingest (proposal edge case 4) bakes the
theme in. Store both (`{"theme": "accent1", "resolved_hex": "#0054A6"}`)
so a future brand-theme change can re-resolve without re-ingesting.

### 5. Data hygiene placement
`template_store/` holds client-branded assets and schemas whose
placeholder text can contain client names. It belongs under `workspace/`
(gitignored) like `workspace/templates/` today — never in the repo. Test
fixtures use synthetic decks only, per root `CLAUDE.md`.

### 6. Relationship to the current mapper: parallel path, not a rewrite
The current mapper (post-Phase-4) is working production tooling with the
COM live preview users like. Template-first should ship as a second,
per-template opt-in pipeline ("Rebuild mode") and earn its way to
default. Do not delete or freeze the current mapper until template-first
has survived a full month-end cycle on real templates.

### 7. Naming collision
The proposal's file layout (`ingestion_engine/templates/ingester.py`)
collides with this app's existing vocabulary, where "ingestion" means
vendor-export parsing. Suggested placement in this codebase:
`app/engine/template_ir/` (schema.py, ingest.py, classify.py, build.py)
with the review/mapping GUI joining `app/mapper/`. Keep "ingest" scoped
as "template ingest" in all naming.

## How it maps onto the existing codebase (reuse, don't rebuild)

| Proposal piece | Existing code to reuse | Notes |
|---|---|---|
| Ingest (shape walk) | `engine/pptx_mapper._scan_with_pptx` | Extend, don't replace — already emits `shape_uid`, types, text |
| Classification review GUI | `mapper/slide_view.py` patterns, `ui/utils.fit_window` | Same Tk idioms; green/orange overlay is new |
| Slot mapping GUI | `mapper/sidebar.py`, `mapper/query_builder.py`, `mapper/format_popup.py` | Metrics catalog, saved queries, format details all carry over |
| Mapping state ownership | `mapper/mapping_model.py` (MappingModel pattern) | New model class, same single-owner + observer discipline |
| Data source → values | `engine/metrics_catalog.get_available_metrics`, `engine/query_resolver` | Slot mappings are structurally the same as assignments+queries |
| Build (text formatting) | `engine/pptx_formats` | Number/date/case rules apply unchanged to slot values |
| Multi-line text writing | `engine/pptx_fill._explode_newlines` | Real `<a:br/>` insertion, already tested |
| Outcome telemetry | `engine/fill_report.FillReport` + `fill_history.jsonl` | A build report is a fill report with a different phase name |
| AI access | `engine/workflow`, `app/cli.py`, `app/mcp_server.py` | Add `ingest_template` / `build_from_template` tools when ready |
| Headless testing | `tests/test_pptx_fill_golden.py` builder helpers | Same synthetic-deck approach |

**Known prerequisite already tracked in STATUS "Noticed":**
`resolve_query()` ignores the query builder's `campaigns`/`sources`/
`values`/`top_n` keys. Template-first makes saved specs re-resolvable by
design, so this resolver-fidelity gap must be fixed *before* slot
mappings can safely reference builder queries.

## Suggested phasing (each phase independently shippable)

- **Phase A — Schema + ingest + verbatim rebuild.** Pydantic schema
  (versioned from day 1), ingester storing parsed schema *plus* original
  shape XML, builder that reconstructs a deck entirely from verbatim
  copies. Acceptance: rebuilt deck is visually identical to the original.
  This proves the fidelity foundation before any mapping exists, and the
  round-trip is fully CI-testable (no COM anywhere in this pipeline).
- **Phase B — Text slots end to end.** Classification heuristics + review
  GUI + slot registry + mapping GUI + dynamic text build. First real
  value: monthly text/KPI decks without the in-place mapper.
- **Phase C — Charts and tables.** Chart-part cloning + data injection;
  table cell mapping. The long tail, isolated so it can't block A/B.
- **Phase D — Re-ingest reconciliation + AI tools.** Template-evolution
  diffs; `ingest_template`/`build_from_template` MCP tools.

**Core test invariant (all phases):** ingest → build → re-ingest must
yield an equivalent schema (round-trip idempotence), and every build
emits a report — silent partial output is the failure mode this project
has been burned by twice (fill placeholders, Excel re-export).

## Library landscape (evaluated 2026-07-15)

Could a different Python PowerPoint library avoid the issues above? No
pure-Python one — and the biggest risk (brand drift from schema
rebuilds) is architectural, not library-bound: it follows from
re-authoring shapes off an abstract description, whichever library
authors them. The verbatim-XML amendment fixes it *within* python-pptx,
which exposes the raw lxml elements that strategy needs.

| Option | What it would solve | Why not (or when) |
|---|---|---|
| python-pptx (current) | Everything in Phases A/B with the verbatim-XML strategy; fully CI-testable | Chart-part cloning is manual plumbing (Phase C); no slide rendering |
| Aspose.Slides for Python (commercial, .NET-based) | First-class high-fidelity cloning; complete chart APIs (collapses Phase C); SmartArt; **renders slides to PNG without PowerPoint** (would delete the COM preview/thumbnail apparatus) | Per-dev license (~$1k+); bundles a .NET runtime into a copy-the-folder portable app; a procurement event that dents the "no new vendors" security pitch. Re-evaluate at Phase C if chart-heavy templates dominate |
| Spire.Presentation (commercial, .NET-based) | Similar shape to Aspose, generally a tier below | Free tier limited to a few slides; same procurement/runtime costs |
| Thin python-pptx wrappers (pptx-template etc.) | Nothing we haven't built | Token replacement over the same engine |
| LibreOffice headless (UNO) | Free rendering/conversion | Different rendering engine — fidelity drift on PowerPoint-authored decks is the norm; disqualifying when brand exactness is the goal |
| PowerPoint COM | Perfect fidelity by definition | The fragility this proposal exists to escape; Windows-only, kills the CI-testable pipeline |

**Decision:** stay on python-pptx; keep `ingest.py`/`build.py` the only
modules that touch the library so the IR isolates a future swap.
Re-evaluate Aspose.Slides specifically at Phase C (charts), priced
against real chart-heavy template volume.

## Open questions — PROVISIONAL answers in effect (2026-07-15)

The owner approved starting Phase A; the question round was cut off by a
tooling failure, so Phase A proceeded on the recommended defaults below.
**Owner: confirm or override each — only #4 gates Phase B's schema.**

1. Fidelity bar → **pixel-exact (provisional)**: Phase A renders static
   shapes exclusively from verbatim XML, so this is satisfied by
   construction.
2. Preview → **static preview of the built deck (provisional)**: no COM
   anywhere in the template-first pipeline.
3. Store location → **`workspace/template_store/<template_id>/`
   (provisional)**: implemented as the default in
   `engine/template_ir/ingest.py`; gitignored.
4. Pivot-driven generation merge → **CONFIRMED by owner 2026-07-16**:
   slot mappings reuse the FULL existing query schema — simple keys AND
   Advanced Query Builder (pivot) queries, resolved through the same
   `resolve_query`/`build_pivot` path the current mapper uses (fidelity
   fixed in v1.32.0). Pivot-driven TABLE/CHART generation stays in
   Phase C. This froze the Phase B mapping schema
   (`engine/template_ir/mapping.py`, mapping.json schema_version 1.0).

## Phase A — DONE 2026-07-15 (all-green in CI)

`app/engine/template_ir/` (schema.py stdlib-dataclass IR + JSON store;
ingest.py deck → IR + verbatim shape XML + extracted image assets keyed
by relationship id; build.py IR → new deck by verbatim XML copies with
image relationships re-targeted). Charts are detected and marked
unsupported — the builder skips them and reports (Phase C). Tests in
`tests/test_template_ir.py` prove the round-trip invariant: ingest →
build → re-ingest is equivalent (XML byte-identical modulo relationship
ids), run formatting and image bytes survive, and skips are always
reported. Next: Phase B (classification + review GUI + slots + mapping).

## Phase B — DONE 2026-07-16 (text slots end to end)

Schema v1.1 (v1.0 stores load unchanged): shapes carry
classification/classify_reason/excluded; the template gains a
`slot_registry` whose slots pin to shape identity + placeholder text
(Critique 3 honored — paragraph/run indices are stored as hints only).
`classify.py` suggests static/dynamic with reasons and names slots from
paired all-caps labels; `mapping.py` maps slots onto the current mapper's
query schema per confirmed open question 4, with map-time type
validation; `build.py` writes slot values into the verbatim copies via
`pptx_fill`'s replace machinery (multi-line → real `<a:br/>`; group
descent) and reports slots_filled/slots_unfilled on every build. Review
GUI (`mapper/template_review.py`) and slot-mapping GUI
(`mapper/slot_mapper.py`, reuses the Advanced Query Builder window via
its wizard duck-type) are thin shells over the tested helpers; entry
points in Settings → Templates and the report template selector. Charts
and tables classify dynamic but take no slots until Phase C; re-ingest
reconciliation (classify_template overwrites review work if re-run) is
Phase D as planned. Tests: `tests/test_template_slots.py`.

## Cross-references

- Proposal: `../proposals/TEMPLATE_FIRST_MAPPER_2026-07-15.md`
- Current-mapper reliability work: `MAPPER_RELIABILITY_ROADMAP_2026-07-12.md`
  (Phases 1–4 done; Phase 5 small fixes still pending and unaffected)
- Authoritative architecture: `../MODEL_HANDOFF.md`, `../CURRENT_ARCHITECTURE.md`
- Session workflow and data hygiene: root `CLAUDE.md`, `STATUS.md`
