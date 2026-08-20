# Handoff Plan Review — 2026-08-20

Adversarial review of the 2026-08-20 planning pass (HANDOFF.md, buildout amendments,
DOCUMENTATION_STANDARDS.md, TESTING_STANDARDS.md, STATUS/CHANGELOG/index updates), run
**before publishing** the pass. Three independent skeptical reviewers, each with a
distinct lens, were instructed to verify every falsifiable claim against the actual code
and tests, rate each part 1–10 honestly, and report defects with severity and a concrete
fix. This report is immutable; dispositions describe the same-pass revision that followed.

## Panel

| Lens | Mandate | Verdict (abridged) |
|---|---|---|
| Architecture / feasibility | Attack DEC-5/7/8, trace `build_from_template` slot consumption, check laws vs `test_architecture.py`, probe the dashboard design | "Unusually well-grounded… NOT safe to commit as-is": 1 blocker (theme/master fidelity), 4 majors |
| Cold-start pickup / completeness | Re-run the pickup guide literally; check every owner ask is covered; verify commands and cross-references | "Not safe to commit as-is, but close": 3 blockers (phantom verb, false `python -m app` claim, phantom review report), 3 majors |
| Process / standards enforceability | Attack the standards' enforceability and self-consistency; verify STATUS honesty against reproduction | "Nearly every falsifiable claim survived… fix the blocker and three majors first": 1 blocker, 4 majors |

Scores per part (P1 framing/lineage, P2 decisions, P3 flow/architecture, P4 phases,
P5 risks/pickup/questions, P6 buildout amendments, P7 doc standards, P8 test standards,
P9 STATUS/CHANGELOG/pointers):

| Part | Architecture | Pickup | Process | Mean |
|---|---|---|---|---|
| P1 | 9 | 9 | 9 | 9.0 |
| P2 | 6 | 7 | 8 | 7.0 |
| P3 | 5 | 6 | 7 | 6.0 |
| P4 | 7 | 7 | 7 | 7.0 |
| P5 | 8 | 6 | 9 | 7.7 |
| P6 | 7 | 7 | 8 | 7.3 |
| P7 | 7 | 7 | 6 | 6.7 |
| P8 | 8 | 9 | 8 | 8.3 |
| P9 | 6 | 5 | 6 | 5.7 |
| **Lens mean** | **7.0** | **7.0** | **7.6** | **7.2** |

What survived attack unbroken (verified independently by 2–3 reviewers): the 313/298/15
test evidence, fork lineage hashes, the two-mapping-systems table, the KPI laws and their
approval dates, the tracked-mappings gitignore gotcha, metric-dictionary contents (20
families / 11 levels), DEC-7's stdlib feasibility, and DEC-5's core claim that
`build_from_template` accepts pre-resolved `slot_values` for all four slot types.

## Blockers (all fixed in this pass)

1. **Template-first theme/master fidelity gap.** `build_from_template` re-emits shapes
   into a bare `Presentation()`; ingest stores no theme/master/layout/background parts, so
   theme-referenced formatting silently falls back to the default Office theme — the exact
   silent brand drift the project fears, and DEC-6 would have had a future session author
   the branded template straight into it. *Fixed:* DEC-5 states the limit and the Stage 4
   entry decision; DEC-6 adds authoring constraints; RSK-12 added; buildout Stage 4 and
   TESTING_STANDARDS §6 require a `schemeClr` golden test.
2. **Phantom `validate_staging` verb.** Presented as an existing workflow verb; it exists
   nowhere (not even a stub), no CLI command exists, and no phase delivered it — while the
   §5 flow and the dashboard depend on it. *Fixed:* marked "added Stage 3" everywhere;
   explicit Phase 3 / buildout Stage 3 deliverable (+ CLI `validate`, contract tests).
3. **False `python -m app` claim, unrecorded real bug.** The pickup guide said it "exits 2
   by design"; it actually crashes (`ModuleNotFoundError: config`, exit 1) because
   `app/__main__.py` lacks the CLI's `sys.path` bootstrap. *Fixed:* pickup guide corrected;
   bug recorded in STATUS.md findings; fix scheduled with the Stage 5 entry-point rework.
4. **STATUS cited this review at a path that did not exist.** The top-authority document
   asserted completed verification with a dead link. *Fixed:* this report is committed at
   that exact path in the same change; STATUS reworded to match reality.

## Majors (all fixed in this pass)

- **Dashboard flow-state had no owner** (`describe_state()` knows nothing of dump/workbook/
  validation; frontend forbidden to own state) → §6.4 now specifies deriving flow state
  from durable artifacts via an extended, documented `describe_state()`; Phase 5 delta.
- **Stage 4 wiring omitted formatting and image payload shape** (raw scalars would ship
  `1234567.0`; `_apply_image` needs a dict) → §6.3 rewritten with the three contract
  details; Q1 extended to the raw-vs-display-string decision.
- **Slot-to-workbook-key contract undefined** (mapping slots are queries over parsed data,
  forbidden in Stage 4) → slot resolution moved to staging time, one row per slot keyed by
  slot name, mapping hash in `_Meta`; Phase 2 delta.
- **`generate_staging` signature misstated in the buildout** (`profile` parameter doesn't
  exist) → verb list corrected to actual stub signatures, current-vs-planned annotated.
- **Positional doc-precedence rule contradicted the buildout's gate authority** → replaced
  with fact-scoped authority (state/gates/decisions/laws each have one owning document).
- **HANDOFF §7 duplicated the buildout's gates verbatim** (violating one-fact-one-home) →
  gates removed from HANDOFF; phases now list only objectives and deltas.
- **Missing EOC requirements document** referenced by the doc index → claim corrected;
  recorded requirements source is HANDOFF §1 + DEC-1…4; noted in STATUS findings.
- **Broken `reviews/` link** in the index → directory exists with this report.

## Minors (17 unique; all adopted except one)

Fixed: T-ARCH-4 label wording (+ STATUS finding to add the label), `template_store`
path-definition footnote, `SCHEMA_VERSION`/Chart-Data consistency (payload sheet is part
of the initial Stage 2 contract; later additions bump), DEC-8 sheet renamed "Chart & Table
Data", DEC-6 gitignore two-layer rules (directory-exclusion semantics + global `*.xlsx`),
DEC-7 token transport specified, DEC-3 variant persistence tied to the bump rule, "or
upload" removed from the flow, hidden developer CLI commands documented, CHANGELOG
headings disambiguated, CI-evidence pointer added to STATUS, loopback carve-out in
TESTING_STANDARDS §4, per-stage minimums aligned to gates and marked as a floor, real
fixture identities cited (+ `*.xlsx` fixture negation note), `workspace/README.md`
un-staled, `AI_CONTEXT.md`/root `README.md` pointers updated, Excel file-lock and
mtime-race handling specified (RSK-11 controls), profile-editor contract surfaced as Q6,
changelog cadence widened, "all machine-enforced" law wording corrected.

Not adopted: creating an empty `documentation/proposals/` directory now — the standards
table marks it "created on first use" instead.

## Outcome

All 5 blockers (4 unique) and 12 unique majors fixed before publishing; the revised pass
is the version committed alongside this report. Residual open items live as STATUS.md
recorded findings and HANDOFF §10 questions Q1–Q6 with named decide-at stages.
