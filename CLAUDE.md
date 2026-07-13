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
   - Make sure the release is covered by the top entry of `documentation/CHANGELOG.md` — the release notes are taken from it.
   - Run the **Release** workflow (`.github/workflows/release.yml`) on the released ref with the version as input; it creates the tag and the GitHub Release. Remote Claude Code sessions cannot push tags or call the release API directly (the git proxy only allows the designated branch), so the workflow is the supported path.
   - Record the version in STATUS.md's decisions log.
   - If the PR is squash-merged, release the merge commit on main (not the branch head) so the tag stays on the default branch's history.
   - Trivial bookkeeping/doc-only PRs (e.g. recording a release in STATUS.md) do not get their own release.

## Git notes for future sessions

- **GitHub's own merge commits are not yours to fix.** After a PR merges and the working branch is reset onto main, the branch tip is the PR merge commit that GitHub created server-side (committer `GitHub <noreply@github.com>`, GPG-signed with GitHub's web-flow key — it shows as **Verified** on GitHub). The stop hook's commit-signature check can misread it as an unverified local commit and suggest `git commit --amend --reset-author`. **Do not amend or rebase it** — rewriting a commit that already exists on main forks the branch's history. Just fast-forward push the branch (`git push -u origin <branch>`) so local and remote match, and only ever reset-author commits you actually authored in the session.
- After a PR for the designated branch merges, restart the branch from main (`git fetch origin main && git checkout -B <branch> origin/main`) before follow-up work; merged PRs are never reused.

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
