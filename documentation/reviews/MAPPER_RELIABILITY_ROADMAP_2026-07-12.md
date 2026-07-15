# Mapper Reliability Roadmap — Status & Next Steps

**Date:** July 12, 2026 (Claude session)
**Suite:** 212 automated tests passing (171 baseline → 198 after golden tests → 212 after tracking)
**Demo rule:** v1.22 stays the frozen NYC demo build. Everything below ships only after their Windows verification passes.

---

## Completed this session

### Phase 1 — Golden-file safety net (DONE, fully verified on Linux)
- `developer/tests/test_pptx_fill_golden.py` (19 tests): fill engine contracts —
  run-formatting preservation, partial/cross-run replacement, legacy mapping
  schema, skip/missing/out-of-range handling, currency/numpy/date formatting,
  image geometry, query resolution incl. failure isolation, text case matching.
- `developer/tests/test_pptx_mapping_store.py` (8 tests): mapping save/load
  round-trip, corrupt-file handling, template listing, scan structure, and
  `test_scan_and_fill_agree_on_shape_identity` — the positional-ID contract
  the Phase 4 refactor will deliberately change.
- Quirks discovered and locked as current behavior (decide later if desired):
  - All-caps placeholder tokens (e.g. `[CLIENT]`) force ALL-CAPS values.
  - "May 1, 2026"-style template text detects as ordinal ("June 1st, 2026").

### Phase 2 — Success/failure tracking (DONE, needs Windows COM pass)
- `app/engine/fill_report.py`: per-fill `FillReport` (filled / images /
  skipped / missing metrics / unmatched placeholders / missing images /
  failed queries / errors) + `fill_history.jsonl` telemetry in workspace/logs.
- `app/engine/pptx_fill.py`: `fill_template_report()` returns (path, report);
  `fill_template()` unchanged for old callers. Unmatched `replace_text`
  placeholders — previously a silent no-op — are now reported.
- `app/engine/pptx_live.py`: all 8 public COM methods health-tracked.
  3 consecutive failures → preview self-disables, fires one `on_disabled`
  notice, and Save & Fill falls back to the tested python-pptx engine
  (existing `is_active()` route). Successes reset the counter.
- UI: mapper Save & Fill and review Auto-Fill dialogs show the report
  summary; "With Gaps" warning variant when `report.ok` is False.

---

## Windows verification checklist (run before merging into demo build)

1. `python -m pytest developer/tests -q` → expect **212 passed**.
2. Normal fill: map a template, Save & Fill → dialog shows
   "Filled N text assignment(s)." and `workspace/logs/fill_history.jsonl`
   gains one line.
3. Gap fill: map one shape to a metric NOT in the export → "Report
   Generated — With Gaps" names the metric; history line has `"ok": false`.
4. Placeholder gap: set a `replace_text` that doesn't exist in the template
   → dialog lists it under "Placeholder text not found".
5. COM death drill: open mapper with live preview running, kill POWERPNT.EXE
   in Task Manager, keep assigning. Expect within ~3 actions: one
   "Live Preview Off" warning, mapper stays usable, Save & Fill still
   produces the report via the built-in engine.
6. Preview health after recovery: reopen mapper fresh → preview works again
   (health state is per-session, nothing persisted).

If any step fails: demo v1.22, report symptoms back, nothing is lost.

---

## Next steps (agreed order, not yet started)

### Phase 3 — MappingModel extraction (~2–3 days, post-NYC)
Single pure-Python owner of mapping state in `app/mapper/`; all mutations go
through it; Tk UI and live COM preview become read-only observers that
re-render from it. Kills the four-way state sync (widgets / shape dicts /
COM working copy / JSON). The golden suite is the acceptance gate: all 212
stay green, UI diffs stay mechanical.

### Phase 4 — Stable shape identity — DONE 2026-07-14 (needs Windows pass)
Shipped as designed, with one deliberate policy reading: **stored identity
that matches nothing → skip + report**, never fall back to the positional
slot (writing there IS the wrong-shape bug). Legacy entries without stored
identity still resolve positionally, bit-for-bit as before.

- Scans emit `shape_uid` (python-pptx `shape.shape_id` / COM `Shape.Id`)
  next to the positional `shape_id`; mapping keys are unchanged.
- Shared pure resolver `engine/shape_identity.py` (uid → unique name →
  positional for legacy entries → None), used by `pptx_fill` and a new
  `PPTXLivePreview._resolve_shape` (fast path: one `Id` read, no
  enumeration when nothing drifted). COM writes no longer "warn but write
  anyway" on drift — they retarget by id or skip.
- `MappingModel.set_scan_identity()` + lazy stamping: mutations stamp
  `shape_uid`/`shape_name` onto the entries they touch; loading/saving old
  mappings changes nothing.
- Fill reports deleted mapped shapes via `FillReport.missing_shapes`
  (drives the "With Gaps" dialog).
- Tests: `test_scan_and_fill_agree_on_shape_identity` replaced by
  reorder/insert/delete drift tests; resolver units in
  `test_shape_identity.py`; COM `_resolve_shape` stub tests in
  `test_live_preview_health.py`. Full suite green (279).

**Windows acceptance checklist (pending):**
1. Full `pytest` pass on Windows.
2. Id parity: scan one template with `_scan_with_pptx` and `_scan_with_com`;
   `shape_uid` must match per shape (load-bearing assumption for the live
   preview; scan+fill are self-consistent either way).
3. Live drift drill: assign a metric, then in the live PowerPoint window
   cut the mapped textbox and paste-in-place (or add a shape above it);
   re-navigate the slide — the value must land on the original shape.
4. Legacy regression: pre-Phase-4 mapping JSON + unchanged template →
   identical fill output; mapping file untouched on disk if nothing edited.
5. Edited-template fill: reorder shapes → values land correctly; delete a
   mapped shape → "With Gaps" reports it, nothing misfilled.
6. Image / chart / table live assignments still work (uid kwarg regression).

### Phase 5 — Small fixes from the July 11 code review — DONE 2026-07-15
- ~~Debounce template-preview selection~~ — done (250 ms debounce in
  review_view's template selector; `_export_thumbnail` was serialized in
  v1.30.0).
- ~~Client wizard drag-select dead code~~ — removed (owner decision: the
  bindings never fired; plain click-to-toggle remains).
- ~~zoomed vs fit_window in ClientWizard~~ — resolved v1.25.0 in favor of
  `fit_window`.
- ~~All-caps case-forcing~~ — owner decision 2026-07-15: KEEP. Values match
  the deck's case ("CLIENT NAME" placeholder → "ACME HOLDING"). Stays
  locked by its characterization tests.
- Bonus (owner decision 2026-07-15): Skip now PRESERVES a shape's
  assignments instead of silently discarding them.

### Parking lot
- **Template-first mapper architecture (v2 direction, accepted 2026-07-15):**
  ingest client decks into a JSON IR with named slots, then BUILD new decks
  instead of editing in place. Supersedes this roadmap's incremental track
  *if/when built* — Phase 5 items above remain worth doing for the current
  mapper in the meantime. Proposal:
  `../proposals/TEMPLATE_FIRST_MAPPER_2026-07-15.md`; critique, phasing, and
  codebase-integration map: `TEMPLATE_FIRST_MAPPER_REVIEW_2026-07-15.md`.
- PySide6 evaluation is a **v2 full-app decision**, not a per-window port
  (two GUI mainloops can't share a process). Revisit only if the tool is
  adopted post-internship.
- `Start Ingestion Engine.bat` runs pip on first launch — do one dry run on
  the actual presentation laptop (network/proxy) before July 14.
