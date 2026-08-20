# Testing Standards

> Internal test rules for Deck Engine. Binding for every stage. Adopted 2026-08-20,
> formalizing the practices the 313-test Stage 0 suite already follows.

## 1. The gate

Tests are the completion gate for every change:

```bash
python -m compileall -q app tests   # always for structural changes
python -m pytest -q                  # must match the count pinned in STATUS.md
```

- **New behavior gets a new test in the same commit.** No exceptions for "obvious" code.
- A stage cannot close with a failing or skipped-without-cause test.
- The expected passing count is pinned in `STATUS.md`; changing it (up or down) is part of
  the diff and gets a sentence explaining why.
- Known environment gap: containers without `tkinter` fail `tests/test_ui_helpers.py` and
  `tests/test_query_builder_helpers.py` at import. CI (ubuntu, Python 3.12) has tkinter;
  treat local Tk-import failures as environmental, everything else as real.

## 2. Required layers (from the buildout, made concrete)

| Layer | What it proves | Existing examples |
|---|---|---|
| Unit | normalization, fingerprints, formatting, path/name safety, KPI math | `test_naming.py`, `test_kpi.py`, `test_pptx_formats.py` |
| Contract | workflow result shapes, staging schema versions, mapping JSON schema round-trips | `test_stage0_harness.py`, `test_pptx_mapping_store.py` |
| Integration | parse → stage → reopen → validate → fill over synthetic files in `tmp_path` | `test_parsers.py`, `test_data_pipeline.py` |
| Golden (semantic) | build a real deck, **reopen it with python-pptx**, assert what a user would see — never binary equality | `test_pptx_fill_golden.py`, `test_template_slots.py` |
| Architecture | the laws: path anchoring, one sanitizer, headless engine, stripped modules stay dead — plus, from Stage 5: engine ⊬ dashboard, dashboard ⊬ tkinter, thin endpoints | `test_architecture.py` |
| Performance | representative 50,000-row export within budget (Stage 1+) | — |
| Manual (Windows) | Office rendering, `os.startfile`, DPI, launchers, file locks, clean-machine install — **only** where automation is not credible; recorded per stage, never assumed | checklist per stage |

Every stage's PR states which layers it touched and which manual checks were **not** run.

## 3. Fixture rules

- **Synthetic only, deterministic always.** No real client data, campaign exports,
  credentials, or internal reports — ever, including "just for debugging". Fixtures are
  generated in-test (see `tests/conftest.py`) or checked in under `tests/fixtures/` with
  obviously fake identities (`FAKE-001`, `Acme Test Co`).
- No randomness without a fixed seed; no time-dependent assertions without freezing time.
- Stage 1 delivers the shared **synthetic fixture factory**; later stages extend it rather
  than inventing parallel fixture styles.

## 4. Isolation rules (what keeps the suite trustworthy)

- All file I/O under pytest `tmp_path`; path constants redirected via `monkeypatch`
  (`paths.TEMPLATES_DIR`, `MAPPINGS_DIR`, …) — never write into the real `workspace/`.
- Module-global caches get autouse resets (pattern: `_reset_dictionary_cache`,
  `_isolate_fingerprints` in `conftest.py`). New global state must ship with its reset.
- No network, no COM, no Office, no display in automated tests. UI helpers are tested
  through fakes (`_FakeRoot`), dashboard endpoints through direct handler calls — no
  browser, no bound port needed for logic tests.

## 5. Behavioral conventions

- **Loud failure is the contract.** Fill/build failures must appear in the
  `FillReport`/findings and must not abort sibling work; tests assert both ("must not
  raise" *and* `report.ok is False`). Silent no-ops may only exist when a test locks them
  deliberately, with a comment saying so.
- **Golden = characterization.** Golden tests lock current user-visible behavior,
  including documented quirks. Intentionally changing behavior means editing the golden
  test in the same commit, with the change called out in the PR.
- **Tests are the decision log's teeth.** Business rules carry their approval date in the
  test docstring (existing pattern: "Approved 2026-07-14"). Do not weaken such a test
  without a new recorded decision.
- Regression tests name the bug they pin (one line: symptom + date), so nobody "cleans
  up" the guard later.

## 6. Per-stage minimums (forward-looking)

- **Stage 1:** fingerprint stability under column reorder; unknown-fingerprint refusal;
  malformed-input error quality; dictionary passthrough notes; 50k-row perf check.
- **Stage 2:** full round-trip (write → close → reopen → read) per supported type; no
  formulas/macros/links in output; unknown `SCHEMA_VERSION` fails loudly; atomic write
  interrupted-write test.
- **Stage 3:** one focused test per blocking rule in the validation catalogue; a
  warnings-never-block test; quarantine placement test.
- **Stage 4:** builder-cannot-see-parsed-data architecture test; workbook → expected deck
  golden test; unmapped/unfilled slot loudness; inherited golden suite untouched.
- **Stage 5:** endpoint thin-shell test; token-required test; localhost-only bind test;
  sanitizer-on-inputs test; full synthetic flow through the HTTP layer; variant panel
  renders from manifest with the flag off by default.
- **Stage 6:** launcher/settings/template-health covered by contract tests; Windows items
  go to the manual checklist.

## 7. CI

CI runs on every push/PR (ubuntu, Python 3.12): install → pytest. Keep it green; a red
main is the top priority. Windows/Office claims are out of CI's reach by definition —
Linux CI **may not** claim Windows acceptance (see `CLAUDE.md`). Release publishing stays
gated on `VERSION` not containing `dev`.
