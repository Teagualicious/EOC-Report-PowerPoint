# Project Template

Starter template for AI-assisted Python projects. Practices are enforced by the files, not by memory:

| File | What it enforces |
|---|---|
| `CLAUDE.md` | Session behavior: read state first, tests before done, minimal code, data hygiene |
| `STATUS.md` | Phase-based handoffs — any fresh session picks up where the last left off |
| `.github/workflows/ci.yml` | pytest runs on every push; red X on anything broken |
| `.gitignore` | Data files and secrets can't be committed by accident |
| `tests/` | pytest scaffold with synthetic fixtures only |

## New project checklist

1. Create repo from this template
2. Edit the "Project overview" section of `CLAUDE.md`
3. Fill in "Next up" in `STATUS.md`
4. Start a Claude Code session and go
