# Deck Engine

Deck Engine is a correctness-first fork of the Jughead Data Engine for a four-step reporting workflow:

1. Select one campaign export.
2. Generate an editable Excel staging workbook.
3. Review and save the workbook in Excel.
4. Build a pre-mapped PowerPoint deck strictly from the saved workbook.

## Current build state

The repository is in **Fork Stage 0**. The upstream v1.36.0 fill, parser, KPI, mapping, and formatting cores are preserved, while the search-workbook, VBA, MCP, multi-window analyst UI, and PowerPoint COM-preview runtime have been removed.

The analyst desktop application intentionally does not launch yet. The available smoke command is:

```bash
python -m app.cli list-templates
```

The complete staged implementation plan is [`documentation/DECK_ENGINE_BUILDOUT.md`](documentation/DECK_ENGINE_BUILDOUT.md). Upstream documentation and the EOC requirements review are retained under [`documentation/upstream/`](documentation/upstream/).

## Development

```bash
python -m pip install -r requirements.txt
pytest
python -m compileall -q app tests
```

No real client data, campaign exports, credentials, or internal reports may be committed. All fixtures must be synthetic.
