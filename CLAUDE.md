# CLAUDE.md

Instructions for Claude Code sessions in this repository. Read STATUS.md before starting any work.

## Project overview

<!-- EDIT PER PROJECT: one paragraph. What this is, who uses it, what "done" looks like. -->
(Describe the project here.)

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
