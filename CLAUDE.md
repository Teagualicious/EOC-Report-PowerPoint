# CLAUDE.md

Instructions for coding sessions in the Deck Engine repository. Read `STATUS.md` before changing anything and use `documentation/DECK_ENGINE_BUILDOUT.md` as the authoritative staged specification.

## Product

Deck Engine is a local-first Windows reporting tool: one campaign dump → editable Excel staging workbook → deterministic PowerPoint deck. It is forked from Jughead Data Engine v1.36.0, preserving its tested parser, KPI, mapping, formatting, and PowerPoint fill cores.

## Stack

- Python 3.11+; standard library before new dependencies
- `pytest` tests in `tests/`
- Runtime dependencies in `app/requirements.txt`
- `openpyxl` for staging; `python-pptx` for fill; no runtime COM/pywin32

## Session workflow

1. Read `STATUS.md`; do not re-derive project state from scratch.
2. Work only on the named stage/task. Record unrelated findings in `STATUS.md`; do not fix them silently.
3. Tests are the gate. Run `pytest`; new behavior gets tests. Run `python -m compileall -q app tests` for structural changes.
4. Close every stage with a clear commit and an updated `STATUS.md` containing what changed, checks run, decisions, and the next entry condition.
5. Publish through a branch and draft PR. Do not push feature work directly to `main`.
6. `VERSION` drives releases only for non-development versions. Values containing `dev` must not publish a release.

## Code and architecture rules

- Minimal code that solves the current stage; no speculative managers/handlers.
- Never remove trust-boundary validation, I/O error handling, or security checks.
- `ui/` and `mapper/` may depend on `engine/`; `engine/` must never depend on Tk/UI.
- All application paths derive from `config.paths`; never use the process CWD for business paths.
- All user-controlled filesystem names go through `config.naming`.
- The Stage 2 staging workbook is the only source of truth for Stage 4 fills. Filling from live parsed data is prohibited.
- Literal staging values only; no runtime COM or formula-recalculation dependency.

## Data hygiene

No real client data, campaign data, credentials, or internal exports may be committed. Generate deterministic synthetic fixtures under `tests/` whenever a realistic shape is needed.

## Windows verification

Office rendering, `os.startfile`, launcher behavior, DPI, file locks, and clean-machine installation are not exercised by Linux CI. Never claim those paths verified without the Windows acceptance checklist for their stage.
