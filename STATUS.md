# STATUS

> The single source of truth for project state. Every session reads this first and updates it last.

## Current phase

Phase 0 — Project setup

## Done

- Repo cleanup: removed committed `tests/__pycache__`, added root `.gitignore` (Python cache/venv, data files csv/xlsx/pptx except `tests/fixtures/*.csv`, secrets), added `.github/workflows/ci.yml` (pytest on push/PR, Python 3.12).
- CI release job: added `release` job to `.github/workflows/ci.yml` — runs only on pushes to `main`, gated on the `test` job passing. Zips repo source (excludes `.git`, `tests`, `__pycache__`) as `{repo-name}_{date}_r{run_number}.zip` and publishes a GitHub Release tagged `build-{run_number}` via `softprops/action-gh-release@v2` using the built-in `GITHUB_TOKEN` (`contents: write`, no new secrets).

## Next up

1. (first task goes here)
2.
3.

## Decisions log

<!-- Date — decision — why. Keeps future sessions from re-litigating settled questions. -->
- YYYY-MM-DD — Repo created from template —
- 2026-07-12 — CI targets Python 3.12 only — matches current toolchain; matrix can be added later if multi-version support is needed.
- 2026-07-12 — Release tag scheme is `build-{run_number}` and zip name embeds date + run number — run_number is monotonic and unique per workflow, giving stable tags without manual versioning.

## Noticed (not yet acted on)

<!-- Problems spotted mid-task but out of scope. Harvest these periodically. -->
- (none)

## How to run

```
pip install -r requirements.txt
pytest
```
