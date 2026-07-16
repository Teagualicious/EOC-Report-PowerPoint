"""Headless terminal interface to the full reporting workflow.

Designed for AI agents and automation, not humans: every command prints a
single JSON object to stdout — ``{"ok": true, "data": ...}`` on success,
``{"ok": false, "error": "..."}`` with a nonzero exit code on failure.
No prompts, no Tk, no interactivity. The interactive app and this CLI
drive the identical engine code, so numbers always agree.

Examples (from the project root):

    python app/cli.py platforms
    python app/cli.py scan --file input/june_export.xlsx
    python app/cli.py campaigns --file "input/june_export.xlsx=Architect"
    python app/cli.py kpis --file "input/june_export.xlsx=Architect" \
        --client "Acme Appliance Co" --start 2026-06-01 --end 2026-06-30
    python app/cli.py export --file "input/june_export.xlsx=Architect" \
        --client "Acme Appliance Co" --start 2026-06-01 --end 2026-06-30
    python app/cli.py fill --file "input/june_export.xlsx=Architect" \
        --client "Acme Appliance Co" --template "Monthly Reporting Template.pptx" \
        --output "output/Acme Appliance Co/June_Report.pptx"
    python app/cli.py query --file "input/june_export.xlsx=Architect" \
        --client "Acme Appliance Co" --metric Impressions --breakdown zip
    python app/cli.py ingest-template --pptx "input/Client Deck.pptx"
    python app/cli.py build-template --file "input/june_export.xlsx=Architect" \
        --client "Acme Appliance Co" --template client_deck \
        --output "output/Acme Appliance Co/June_Report.pptx"

``--file`` may repeat and uses ``path=Platform`` (the platform must be
configured in the app first). ``--campaigns`` limits a client to specific
campaigns (comma-separated); omitted means all campaigns in the files.
"""

import argparse
import json
import logging
import os
import sys

# Same bootstrap as app/main.py: make app/ modules importable when the CLI
# is invoked as `python app/cli.py` from the project root.
_APP_DIR = os.path.dirname(os.path.abspath(__file__))
if _APP_DIR not in sys.path:
    sys.path.insert(0, _APP_DIR)


def _parse_file_args(file_args):
    pairs = []
    for spec in file_args or []:
        if "=" not in spec:
            raise ValueError(
                f"--file needs path=Platform, got: {spec}")
        path, platform = spec.rsplit("=", 1)
        pairs.append((path.strip(), platform.strip()))
    return pairs


def _parsed_or_fail(args):
    from engine import workflow
    pairs = _parse_file_args(args.file)
    if not pairs:
        raise ValueError("At least one --file path=Platform is required")
    parsed, failures = workflow.parse_files(pairs)
    if not parsed:
        detail = "; ".join(f"{n}: {r}" for n, r in failures) or "no files"
        raise ValueError(f"No files could be parsed — {detail}")
    return parsed, [{"file": n, "reason": r} for n, r in failures]


def _client_data(args, parsed):
    from engine import workflow
    campaigns = ([c.strip() for c in args.campaigns.split(",") if c.strip()]
                 if getattr(args, "campaigns", None) else None)
    return workflow.client_dataset(parsed, args.client, campaigns,
                                   args.start or "", args.end or "")


def cmd_platforms(_args):
    from engine import workflow
    return workflow.list_platforms()


def cmd_templates(_args):
    from engine import workflow
    return workflow.list_templates()


def cmd_scan(args):
    from engine import workflow
    return workflow.scan_export(args.file[0].split("=", 1)[0]
                                if "=" in args.file[0] else args.file[0])


def cmd_campaigns(args):
    from engine import workflow
    parsed, failures = _parsed_or_fail(args)
    return {"campaigns": workflow.list_campaigns(parsed),
            "failures": failures}


def cmd_kpis(args):
    from engine import workflow
    parsed, failures = _parsed_or_fail(args)
    result = workflow.kpis_for(_client_data(args, parsed))
    result["failures"] = failures
    return result


def cmd_export(args):
    from engine import workflow
    parsed, failures = _parsed_or_fail(args)
    client_data = _client_data(args, parsed)
    result = workflow.export_workbook(client_data, args.client,
                                      args.output_dir or "",
                                      inject_vba=not args.no_vba)
    result["kpis"] = workflow.kpis_for(client_data)
    result["failures"] = failures
    return result


def cmd_fill(args):
    from engine import workflow
    parsed, failures = _parsed_or_fail(args)
    result = workflow.fill_deck(_client_data(args, parsed), args.template,
                                args.output, args.client,
                                args.start or "", args.end or "")
    result["failures"] = failures
    return result


def cmd_query(args):
    from engine import workflow
    parsed, failures = _parsed_or_fail(args)
    result = workflow.query_metric(
        _client_data(args, parsed), args.metric, args.breakdown, args.agg,
        args.filter, args.client, args.start or "", args.end or "")
    result["failures"] = failures
    return result


def cmd_ingest_template(args):
    from engine import workflow
    return workflow.ingest_template_store(args.pptx, args.template_id)


def cmd_template_stores(_args):
    from engine import workflow
    return workflow.list_template_stores()


def cmd_build_template(args):
    from engine import workflow
    parsed, failures = _parsed_or_fail(args)
    result = workflow.build_template_report(
        _client_data(args, parsed), args.template, args.output, args.client,
        args.start or "", args.end or "")
    result["failures"] = failures
    return result


def _add_file_args(p, client_required=True):
    p.add_argument("--file", action="append", required=True,
                   metavar="PATH=PLATFORM",
                   help="export file and its configured platform (repeatable)")
    p.add_argument("--client", required=client_required,
                   help="client name for the report")
    p.add_argument("--campaigns", help="comma-separated campaign names "
                                       "(default: all found)")
    p.add_argument("--start", help="reporting period start (YYYY-MM-DD)")
    p.add_argument("--end", help="reporting period end (YYYY-MM-DD)")


def build_parser():
    parser = argparse.ArgumentParser(
        prog="ingestion-engine",
        description="Headless JSON interface to the reporting workflow.")
    sub = parser.add_subparsers(dest="command", required=True)

    sub.add_parser("platforms", help="list configured platforms"
                   ).set_defaults(fn=cmd_platforms)
    sub.add_parser("templates", help="list saved PowerPoint templates"
                   ).set_defaults(fn=cmd_templates)

    p = sub.add_parser("scan", help="inspect an export's structure")
    p.add_argument("--file", action="append", required=True)
    p.set_defaults(fn=cmd_scan)

    p = sub.add_parser("campaigns", help="list campaigns in the files")
    _add_file_args(p, client_required=False)
    p.set_defaults(fn=cmd_campaigns)

    p = sub.add_parser("kpis", help="compute review KPIs and data flags")
    _add_file_args(p)
    p.set_defaults(fn=cmd_kpis)

    p = sub.add_parser("export", help="write the unified Excel workbook")
    _add_file_args(p)
    p.add_argument("--output-dir", help="output root (default: app setting)")
    p.add_argument("--no-vba", action="store_true",
                   help="skip the Windows Excel search injection")
    p.set_defaults(fn=cmd_export)

    p = sub.add_parser("fill", help="fill a mapped PowerPoint template")
    _add_file_args(p)
    p.add_argument("--template", required=True,
                   help="template filename in workspace/templates")
    p.add_argument("--output", required=True, help="output .pptx path")
    p.set_defaults(fn=cmd_fill)

    p = sub.add_parser("query", help="resolve a metric query")
    _add_file_args(p)
    p.add_argument("--metric", required=True)
    p.add_argument("--breakdown", default="all",
                   help="breakdown type (zip, network, ...) or 'all'")
    p.add_argument("--agg", default="sum",
                   choices=["sum", "avg", "max", "min", "count"])
    p.add_argument("--filter", default="all",
                   help="campaign name or level value to filter on")
    p.set_defaults(fn=cmd_query)

    p = sub.add_parser("ingest-template",
                       help="ingest (or re-ingest) a deck into the "
                            "template-first store; re-ingest returns the "
                            "review deltas")
    p.add_argument("--pptx", required=True, help="path to the client .pptx")
    p.add_argument("--template-id",
                   help="store id (default: derived from the filename)")
    p.set_defaults(fn=cmd_ingest_template)

    sub.add_parser("template-stores",
                   help="list ingested template-first stores"
                   ).set_defaults(fn=cmd_template_stores)

    p = sub.add_parser("build-template",
                       help="build a deck from a mapped template-first store")
    _add_file_args(p)
    p.add_argument("--template", required=True,
                   help="template store id or directory path")
    p.add_argument("--output", required=True, help="output .pptx path")
    p.set_defaults(fn=cmd_build_template)
    return parser


def main(argv=None):
    logging.basicConfig(level=logging.WARNING)  # stdout stays pure JSON
    args = build_parser().parse_args(argv)
    try:
        data = args.fn(args)
    except Exception as exc:
        logger = logging.getLogger(__name__)
        logger.debug("CLI command failed", exc_info=True)
        print(json.dumps({"ok": False, "error": str(exc)}, default=str))
        return 1
    print(json.dumps({"ok": True, "data": data}, default=str))
    return 0


if __name__ == "__main__":
    sys.exit(main())
