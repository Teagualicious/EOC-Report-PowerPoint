"""MCP server: exposes the reporting workflow as native tools for Claude.

Runs locally over stdio — campaign data never leaves this machine; Claude
sends tool calls and receives only the results. The tools drive the exact
engine code behind the desktop app (engine/workflow.py), so numbers always
match what the UI shows.

Requires the ``mcp`` package (not part of the app's runtime requirements —
end users of the desktop app don't need it):

    python -m pip install mcp

Claude Desktop config (Settings > Developer > Edit Config), adjust paths:

    {
      "mcpServers": {
        "ingestion-engine": {
          "command": "python",
          "args": ["C:\\\\path\\\\to\\\\Jughead-Data-Engine\\\\app\\\\mcp_server.py"]
        }
      }
    }

Read-oriented tools are always on. The two tools that WRITE deliverables
(export_workbook, fill_template) are enabled by default but can be disabled
by starting the server with INGESTION_MCP_READ_ONLY=1 in the environment.
"""

import os
import sys

_APP_DIR = os.path.dirname(os.path.abspath(__file__))
if _APP_DIR not in sys.path:
    sys.path.insert(0, _APP_DIR)

try:
    from mcp.server.fastmcp import FastMCP
except ImportError:  # pragma: no cover - exercised only without the package
    print("The 'mcp' package is required for the AI server:\n"
          "    python -m pip install mcp", file=sys.stderr)
    sys.exit(1)

from engine import workflow

READ_ONLY = os.environ.get("INGESTION_MCP_READ_ONLY", "") == "1"

mcp = FastMCP(
    "ingestion-engine",
    instructions=(
        "Local reporting engine for Spectrum Reach campaign exports. "
        "Workflow: list_platforms -> (scan_export for new layouts) -> "
        "list_campaigns(files) -> get_kpis / query_metric to analyze, "
        "export_workbook for the searchable Excel report, fill_template "
        "for the mapped PowerPoint deck. Template-first pipeline: "
        "ingest_template ingests a client deck into a reviewable store "
        "(re-ingest returns review deltas), list_template_stores shows "
        "stores, build_from_template builds a fresh deck from a mapped "
        "store. Files are given as [{path, platform}] where platform "
        "names a saved app config."),
)


def _pairs(files):
    return [(f["path"], f["platform"]) for f in files]


def _dataset(files, client, campaigns=None, start_date="", end_date=""):
    parsed, failures = workflow.parse_files(_pairs(files))
    if not parsed:
        detail = "; ".join(f"{n}: {r}" for n, r in failures) or "no files"
        raise ValueError(f"No files could be parsed — {detail}")
    data = workflow.client_dataset(parsed, client, campaigns or None,
                                   start_date, end_date)
    return data, [{"file": n, "reason": r} for n, r in failures]


@mcp.tool()
def list_platforms() -> list:
    """List the platforms configured in the app (needed for file parsing)."""
    return workflow.list_platforms()


@mcp.tool()
def list_templates() -> list:
    """List saved PowerPoint templates and whether each has a mapping."""
    return workflow.list_templates()


@mcp.tool()
def scan_export(path: str) -> list:
    """Inspect an export file's structure (sheets, headers, sample row)
    before a platform config exists for it."""
    return workflow.scan_export(path)


@mcp.tool()
def list_template_stores() -> list:
    """List ingested template-first stores (id, dir, slot count, whether
    a slot mapping is saved)."""
    return workflow.list_template_stores()


@mcp.tool()
def list_campaigns(files: list) -> dict:
    """List campaigns found in export files.

    files: [{"path": str, "platform": str}] — platform must be configured.
    """
    parsed, failures = workflow.parse_files(_pairs(files))
    return {"campaigns": workflow.list_campaigns(parsed),
            "failures": [{"file": n, "reason": r} for n, r in failures]}


@mcp.tool()
def get_kpis(files: list, client: str, campaigns: list = None,
             start_date: str = "", end_date: str = "") -> dict:
    """Compute the review-screen KPIs (totals, per-campaign detail, data
    quality flags) for a client from export files."""
    data, failures = _dataset(files, client, campaigns, start_date, end_date)
    result = workflow.kpis_for(data)
    result["failures"] = failures
    return result


@mcp.tool()
def query_metric(files: list, client: str, metric: str,
                 breakdown: str = "all", agg: str = "sum",
                 metric_filter: str = "all", campaigns: list = None,
                 start_date: str = "", end_date: str = "") -> dict:
    """Resolve an advanced metric query (e.g. Impressions by zip, sum) —
    the same engine mapped decks use, so answers match reports."""
    data, failures = _dataset(files, client, campaigns, start_date, end_date)
    result = workflow.query_metric(data, metric, breakdown, agg,
                                   metric_filter, client, start_date,
                                   end_date)
    result["failures"] = failures
    return result


if not READ_ONLY:
    @mcp.tool()
    def export_workbook(files: list, client: str, campaigns: list = None,
                        start_date: str = "", end_date: str = "",
                        output_dir: str = "", inject_vba: bool = True) -> dict:
        """Write the client's unified Excel workbook (Search dashboard +
        Unified Data), merging with any previous export of the same period."""
        data, failures = _dataset(files, client, campaigns,
                                  start_date, end_date)
        result = workflow.export_workbook(data, client, output_dir,
                                          inject_vba=inject_vba)
        result["kpis"] = workflow.kpis_for(data)
        result["failures"] = failures
        return result

    @mcp.tool()
    def fill_template(files: list, client: str, template: str,
                      output_path: str, campaigns: list = None,
                      start_date: str = "", end_date: str = "") -> dict:
        """Fill a mapped PowerPoint template with the client's data and
        return the fill report (what filled, what was missing)."""
        data, failures = _dataset(files, client, campaigns,
                                  start_date, end_date)
        result = workflow.fill_deck(data, template, output_path, client,
                                    start_date, end_date)
        result["failures"] = failures
        return result

    @mcp.tool()
    def ingest_template(pptx_path: str, template_id: str = "") -> dict:
        """Ingest a client deck into the template-first store: shapes are
        auto-classified static/dynamic and placeholder runs become named
        slots (a human reviews in the app). Re-ingesting an updated deck
        carries the review forward and returns the deltas."""
        return workflow.ingest_template_store(pptx_path,
                                              template_id or None)

    @mcp.tool()
    def build_from_template(files: list, client: str, template: str,
                            output_path: str, campaigns: list = None,
                            start_date: str = "", end_date: str = "") -> dict:
        """Build a fresh deck from a mapped template-first store (verbatim
        branding, slot values resolved from the client's data; charts and
        tables injected). Returns the build report — slots filled and
        unfilled, shapes skipped, notes."""
        data, failures = _dataset(files, client, campaigns,
                                  start_date, end_date)
        result = workflow.build_template_report(data, template, output_path,
                                                client, start_date, end_date)
        result["failures"] = failures
        return result


if __name__ == "__main__":
    mcp.run()
