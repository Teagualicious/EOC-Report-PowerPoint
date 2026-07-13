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

### Phase 4 — Stable shape identity (after Phase 3)
Key mappings by PowerPoint `shape.Id` (stable) with name fallback, instead
of positional index. Backward compatible: keep positional keys in the JSON,
add `shape_uid` alongside; fill prefers uid, falls back to index. Update
`test_scan_and_fill_agree_on_shape_identity` deliberately; everything else
must pass unchanged. Removes the "SHAPE INDEX DRIFT" class of wrong-shape
assignments entirely (pptx_live already warns about drift; this fixes it).

### Phase 5 — Small fixes from the July 11 code review
- Debounce template-preview selection (`after(300)`) + serialize
  `_export_thumbnail` per path (PowerPoint process pile-up / PNG race).
- Client wizard drag-select is dead code (Tk implicit grab eats `<Enter>`
  during button press); reimplement via `<B1-Motion>` + `winfo_containing`,
  or remove the drag bindings.
- Decide: `state("zoomed")` vs `fit_window` in ClientWizard (currently both;
  zoom wins).
- Decide: keep or fix all-caps placeholder case-forcing (Phase 1 quirk).

### Parking lot
- PySide6 evaluation is a **v2 full-app decision**, not a per-window port
  (two GUI mainloops can't share a process). Revisit only if the tool is
  adopted post-internship.
- `Start Ingestion Engine.bat` runs pip on first launch — do one dry run on
  the actual presentation laptop (network/proxy) before July 14.
