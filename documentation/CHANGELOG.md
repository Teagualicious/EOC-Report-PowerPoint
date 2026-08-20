# Changelog

## 0.1.0-dev — Planning pass: handoff, standards, dashboard direction (2026-08-20)

- Added `documentation/HANDOFF.md` (context, architecture map, decisions DEC-1…DEC-8, phase plan, pickup guide).
- Amended `documentation/DECK_ENGINE_BUILDOUT.md`: Stage 5 analyst UI is now a localhost web dashboard (replacing the planned Tkinter window); Stage 4 gains the interim Spectrum-branded default template and literal chart-data rule; Stage 6 launcher starts the dashboard.
- Added binding `documentation/DOCUMENTATION_STANDARDS.md` and `documentation/TESTING_STANDARDS.md`; updated the documentation index and `CLAUDE.md` reading order.
- No application or test code changed; survivor baseline remains 313 tests.

## 0.1.0-dev — Fork Stage 0: surgery and architecture harness

- Established the exact Jughead v1.36.0 fork baseline and recorded its 367-test result.
- Removed the MCP server, searchable Excel/VBA export path, PowerPoint live COM preview, and inherited multi-window analyst workflow.
- Preserved the parser, Unified Data, KPI, mapping, formatting, template, and golden PowerPoint fill cores.
- Added the stable `parse_dump`, `generate_staging`, and `build_deck` workflow boundaries plus staging, validation, and campaign-dictionary module contracts.
- Added project-anchored staging/dictionary/quarantine paths, light-only theming, Deck Engine logging, a developer-only mapper launcher, and architecture enforcement tests.
- Development versions are explicitly excluded from automatic GitHub releases.
