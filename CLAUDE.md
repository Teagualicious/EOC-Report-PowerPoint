# CLAUDE.md

Instructions for Claude Code sessions in this repository. Read STATUS.md before starting any work.

## Project overview

Spectrum Reach Reporting Ingestion Engine — a local-first Windows desktop app (Tkinter) that converts advertising-platform exports (CSV/XLSX/XLSM/HTML) into searchable Excel workbooks and fills mapped PowerPoint report templates. Application code lives in `app/`, runtime state in `workspace/`; end users launch via `Start Ingestion Engine.bat`.

Before changing application code, read `AI_CONTEXT.md` (design rules, business invariants, pitfalls) and `documentation/MODEL_HANDOFF.md` (authoritative architecture handoff). The Windows-only Office COM paths (Excel VBA injection, PowerPoint live preview) are not exercised by the automated suite — never claim them verified without a Windows/Office acceptance pass.

## Stack

- Python 3.11+ (standard library preferred over new dependencies)
- Tests: pytest, in `tests/`
- Dependencies: `requirements.txt` — do not add a dependency without noting why in the commit message

## Workflow rules

1. **Start of session:** read STATUS.md to learn current phase, what's done, and what's next. Do not re-derive project state from scratch.
2. **Scope:** work only on the task given. If you notice unrelated problems, list them in STATUS.md under "Noticed" — do not fix them unprompted.
3. **Tests are the gate.** Run `pytest` before declaring any task complete. A task with failing tests is not done. New behavior gets a new test.
4. **End of session (every time):**
   - Update STATUS.md: what changed, what's next, any decisions made
   - Commit with a clear message
   - Leave the repo in a state a fresh session can pick up with zero conversation context
5. **Release after a PR.** When a PR's work is complete, create a release for it:
   - Bump the version (the project's line continues from the inherited v1.22 demo build; bump the minor version for a normal batch of work, e.g. v1.23.0).
   - Create an annotated git tag on the released commit whose message mirrors the CHANGELOG entry, and push the tag: `git tag -a vX.Y.0 -m "..." && git push origin vX.Y.0`.
   - Record the version in STATUS.md.
   - Note: remote Claude Code sessions cannot create GitHub Release pages (no `gh`/API access — tags only); promote the tag to a Release in the GitHub UI if release notes should be user-visible. If the PR is squash-merged, re-tag the merge commit on main so the release stays on the default branch's history.

## Code style

- Minimal code that solves the stated problem. Reuse existing functions before writing new ones. Stdlib before dependencies. One line if one line works.
- Never cut: input validation at trust boundaries, error handling around I/O, anything security-relevant.
- No speculative abstractions. No "manager" or "handler" classes for things that happen once.
- Match the existing style of the file being edited.

## Data hygiene (non-negotiable)

- No real client data, campaign data, credentials, or company-internal exports in this repo. Ever.
- Test fixtures use synthetic data only (see `tests/fixtures/`).
- If a task requires realistic data shapes, generate fake data matching the schema.

## Phase discipline

Work is organized in phases (see STATUS.md). A phase ends with: tests passing, STATUS.md updated, changes committed. Prefer finishing a phase over starting the next one.
