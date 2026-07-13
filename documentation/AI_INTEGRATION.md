# AI Integration Guide

The Ingestion Engine is AI-native: the complete reporting workflow — parse
exports, compute KPIs, build the searchable Excel workbook, fill mapped
PowerPoint decks, answer metric questions — runs headlessly through two
machine-facing interfaces built on the exact engine code behind the desktop
app. Numbers always match what the UI shows, and the automated test suite
covers both paths.

**Everything runs locally.** Claude (or any agent) sends tool calls; the
engine executes on this machine; campaign data never leaves it. Only the
results that would appear on screen enter the conversation.

## The three layers

```text
Desktop app (Tkinter)  ─┐
Terminal CLI (app/cli.py) ─┼──▶  engine/workflow.py  ──▶  parsers / KPI /
MCP server (app/mcp_server.py) ─┘   (one shared service)     Excel / PowerPoint
```

`engine/workflow.py` is the single headless service; the CLI and MCP server
are thin shells over it (enforced by a test). Anything the app can produce,
an AI can produce.

## Option 1 — MCP server for Claude Desktop / Claude Code (recommended)

One-time setup on the reporting machine:

1. `python -m pip install mcp`
2. Claude Desktop → Settings → Developer → Edit Config, add (adjust path):

```json
{
  "mcpServers": {
    "ingestion-engine": {
      "command": "python",
      "args": ["C:\\path\\to\\Jughead-Data-Engine\\app\\mcp_server.py"]
    }
  }
}
```

3. Restart Claude Desktop — the tools appear automatically.

Tools exposed:

| Tool | What it does |
|---|---|
| `list_platforms` | Configured platforms (needed to parse files) |
| `list_templates` | Saved PowerPoint templates + mapping status |
| `scan_export` | Inspect a new export's sheets/columns/sample row |
| `list_campaigns` | Campaigns found in export files |
| `get_kpis` | Review-screen KPIs, per-campaign detail, data flags |
| `query_metric` | Advanced metric query (breakdowns, filters, agg) |
| `export_workbook` | Write the unified Search/Data Excel workbook |
| `fill_template` | Fill a mapped PowerPoint deck, returns the fill report |

Set the environment variable `INGESTION_MCP_READ_ONLY=1` on the server to
disable the two writing tools (analysis-only mode for wider rollouts).

Example prompts once connected:

- *"Scan input/june_export.xlsx and tell me if our Architect platform
  config covers all its columns."*
- *"Using input/june_export.xlsx as Architect, what were Famous Tate's
  impressions by zip in June? Any data flags?"*
- *"Export the June workbook for Famous Tate and fill the Monthly
  Reporting Template — tell me if anything didn't fill."*

## Option 2 — Terminal CLI (any agent, scripts, schedulers)

Every command prints one JSON object (`{"ok": true, "data": ...}` or
`{"ok": false, "error": ...}` with exit code 1). No prompts, no UI.

```bash
python app/cli.py platforms
python app/cli.py scan --file input/june_export.xlsx
python app/cli.py campaigns --file "input/june_export.xlsx=Architect"
python app/cli.py kpis --file "input/june_export.xlsx=Architect" \
    --client "Famous Tate" --start 2026-06-01 --end 2026-06-30
python app/cli.py export --file "input/june_export.xlsx=Architect" \
    --client "Famous Tate" --start 2026-06-01 --end 2026-06-30
python app/cli.py fill --file "input/june_export.xlsx=Architect" \
    --client "Famous Tate" --template "Monthly Reporting Template.pptx" \
    --output "output/Famous Tate/June_Report.pptx"
python app/cli.py query --file "input/june_export.xlsx=Architect" \
    --client "Famous Tate" --metric Impressions --breakdown zip
```

`--file` repeats and takes `path=Platform`; `--campaigns "A,B"` limits a
client to specific campaigns (default: all found in the files).

## What stays human

Template *mapping* (assigning metrics to slide shapes with the live
PowerPoint preview) remains an interactive task in the desktop app. AI
interfaces consume saved mappings; they do not create them. Windows-only
Office integration keeps its usual graceful degradation: `export_workbook`
produces a plain .xlsx with a warning when Excel/VBA injection is
unavailable, and `fill_template` always uses the built-in python-pptx
engine.

## Governance notes

- Query results and fill reports enter the AI conversation; raw exports and
  workbooks do not. Confirm client-data policy for AI tools before rollout.
- The `mcp` package is intentionally NOT in the app's runtime requirements —
  desktop users install nothing extra; only the machine hosting the AI
  integration needs it.
- Every fill is still recorded to `workspace/logs/fill_history.jsonl`,
  whether a human or an AI triggered it.
