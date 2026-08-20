# Documentation Standards

> Rules for writing and maintaining documentation in this repository so that any future
> session — human or AI — can pick the project up cold and act safely. These rules are
> binding for every stage. Adopted 2026-08-20.

## 1. The document hierarchy (one fact, one home)

| Document | Role | Update cadence |
|---|---|---|
| `STATUS.md` | **Single source of truth for project state.** Read first, updated last, every session. | Every session close |
| `documentation/HANDOFF.md` | Context, architecture map, phase plan, decisions. | When architecture or plan changes |
| `documentation/DECK_ENGINE_BUILDOUT.md` | Stage gates, controls, validation catalogue. | Only via a recorded decision |
| `CLAUDE.md` | Session law: workflow, stack, hard rules. | Rarely; keep short |
| `documentation/CHANGELOG.md` | What shipped, per version, removals named as prominently as additions. | Every stage close |
| `documentation/TESTING_STANDARDS.md` / this file | How to test / how to document. | Via recorded decision |
| `documentation/reviews/` | Dated, immutable review reports (`NAME_YYYY-MM-DD.md`). | Append-only |
| `documentation/proposals/` | Accepted-but-unscheduled designs, read with their paired review. | Append-only |
| `documentation/upstream/` | Archived donor docs. **Context, not authority.** | Never (frozen) |

Never state the same fact authoritatively in two places — link to its home instead. If two
documents disagree, the one higher in this table wins, and fixing the disagreement is part
of the change that exposed it.

## 2. AI-compatibility rules (make docs machine-actionable)

1. **Declare the reading order.** `CLAUDE.md` names what to read first; every doc's header
   links its companions. A session must never need tribal knowledge to find the next file.
2. **Exact paths, exact commands.** File references are repo-relative paths in backticks
   (`app/engine/staging.py`). Runnable commands sit in fenced code blocks, copy-pasteable,
   with expected output stated when it matters ("expect the count pinned in STATUS.md").
3. **Pin verifiable numbers, and say how to verify.** Test counts, schema versions,
   baselines — always paired with the command that reproduces them. A stale pinned number
   is a bug; fix it in the same commit that changes it.
4. **State over narrative.** Prefer tables, contracts, and state machines to prose
   history. Where history matters, date it.
5. **Decisions carry IDs and dates.** Product/engineering decisions get a `DEC-n` entry in
   `HANDOFF.md` §3–4 and a dated line in `STATUS.md`'s decisions log. Code comments cite
   them by date (existing convention: "Approved 2026-07-14"). A decision is reversible
   only by a new dated entry — never by silent edits.
6. **Mark what is dead or historical.** Removed features, archived docs, and stubbed
   modules are labeled as such where they're mentioned. Never let a doc describe removed
   behavior as current (the upstream docs are the cautionary example).
7. **Contracts before behavior.** When a stage introduces a data shape (workbook sheet,
   JSON schema, API response), document the contract — fields, types, version, failure
   mode for unknown versions — in the module docstring and reference it from `HANDOFF.md`.
8. **No screenshots as sole evidence.** Diagrams are ASCII/text in the doc so they diff,
   grep, and survive tooling changes. Images may supplement, never replace.
9. **Plain unambiguous language.** No unexpanded acronyms on first use, no "it/this"
   referring across paragraphs, no relative dates ("recently") — absolute dates only.
10. **Entry conditions end every unit of work.** A stage close writes the next stage's
    entry condition into `STATUS.md`, so the next session starts with zero re-derivation.

## 3. Session obligations (docs are part of the diff)

- Every PR that changes behavior updates: the relevant contract docstrings, `STATUS.md`,
  and — if a shipped surface changed — `documentation/CHANGELOG.md`. Reviewers treat a
  missing doc update like a missing test.
- New modules start with a docstring stating: purpose, the stage that owns them, their
  public API, and the laws they must obey.
- Findings outside the current stage are recorded in `STATUS.md` (or as an open question
  in `HANDOFF.md` §10 with a decide-by stage), never fixed silently.
- Never claim Windows/Office behavior verified in any document without the Windows
  acceptance checklist run recorded for that stage.

## 4. Style

- Markdown, ATX headings (`##`), stable heading names (other docs deep-link to them).
- Line width ≤ 100 where practical; tables kept narrow enough to read in a terminal.
- Checklists use `- [ ]` items so completion is diffable.
- File names: current docs `SCREAMING_SNAKE.md`; dated artifacts `NAME_YYYY-MM-DD.md`.
