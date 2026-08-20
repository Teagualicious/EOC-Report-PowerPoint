#!/usr/bin/env python3
"""No-prompt JSON command line interface for Deck Engine."""

from __future__ import annotations

import argparse
import json
import logging
import os
import sys

# Support both ``python -m app.cli`` and ``python app/cli.py``.
_APP_DIR = os.path.dirname(os.path.abspath(__file__))
if _APP_DIR not in sys.path:
    sys.path.insert(0, _APP_DIR)


def cmd_list_templates(_args):
    from engine import workflow
    return workflow.list_templates()


def cmd_state(_args):
    from engine import workflow
    return workflow.describe_state()


def cmd_settings(args):
    from engine import workflow
    if args.set:
        updates = {}
        for item in args.set:
            if "=" not in item:
                raise ValueError(f"--set needs KEY=VALUE, got: {item}")
            key, raw = item.split("=", 1)
            try:
                value = json.loads(raw)
            except json.JSONDecodeError:
                value = raw
            updates[key.strip()] = value
        return workflow.set_settings(updates)
    return workflow.get_settings()


def cmd_parse(args):
    from engine import workflow
    return workflow.parse_dump(args.dump)


def cmd_stage(args):
    from engine import workflow
    return workflow.generate_staging(args.dump, args.template)


def cmd_build(args):
    from engine import workflow
    return workflow.build_deck(args.staging)



def cmd_ingest_template(args):
    from engine import workflow
    return workflow.ingest_template_store(args.pptx, args.template_id)


def cmd_template_stores(_args):
    from engine import workflow
    return workflow.list_template_stores()

def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="deck",
        description="Headless JSON interface to the Deck Engine workflow.")
    sub = parser.add_subparsers(dest="command", required=True)

    sub.add_parser("list-templates", aliases=["templates"],
                   help="list saved PowerPoint templates").set_defaults(
                       fn=cmd_list_templates)
    sub.add_parser("state", help="show paths, version, and current readiness"
                   ).set_defaults(fn=cmd_state)

    settings = sub.add_parser("settings", help="read or update settings")
    settings.add_argument("--set", action="append", metavar="KEY=JSON_VALUE")
    settings.set_defaults(fn=cmd_settings)

    parse = sub.add_parser("parse", help="parse a campaign export")
    parse.add_argument("dump", help="CSV/XLSX/XLSM/HTML export path")
    parse.set_defaults(fn=cmd_parse)

    stage = sub.add_parser("stage", help="generate a staging workbook")
    stage.add_argument("dump", help="campaign export path")
    stage.add_argument("--template", help="mapped template filename")
    stage.set_defaults(fn=cmd_stage)

    build = sub.add_parser("build", help="build a deck from a staging workbook")
    build.add_argument("--staging", help="staging workbook path; default setting if omitted")
    build.set_defaults(fn=cmd_build)

    # Developer-only template authoring commands retained for the mapper.
    ingest = sub.add_parser("ingest-template", help=argparse.SUPPRESS)
    ingest.add_argument("--pptx", required=True)
    ingest.add_argument("--template-id")
    ingest.set_defaults(fn=cmd_ingest_template)
    sub.add_parser("template-stores", help=argparse.SUPPRESS).set_defaults(
        fn=cmd_template_stores)
    return parser


def main(argv=None) -> int:
    logging.basicConfig(level=logging.WARNING)
    args = build_parser().parse_args(argv)
    try:
        data = args.fn(args)
    except Exception as exc:
        logging.getLogger(__name__).debug("CLI command failed", exc_info=True)
        print(json.dumps({
            "ok": False,
            "error": str(exc),
            "user_message": getattr(exc, "user_message", str(exc)),
            "code": getattr(exc, "code", None),
        }, default=str))
        return 1
    print(json.dumps({"ok": True, "data": data}, default=str))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
