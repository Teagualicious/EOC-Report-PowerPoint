"""Outcome tracking for template fills.

A ``FillReport`` records what actually happened during ``fill_template``:
how many assignments landed, which metrics were missing, which
placeholders never matched, which images or queries failed. The mapper
UI shows the summary after every fill, and each report is appended to
``workspace/logs/fill_history.jsonl`` so success/failure is trackable
across sessions.
"""

import json
import logging
import os
from datetime import datetime, timezone

logger = logging.getLogger(__name__)

FILL_HISTORY_NAME = "fill_history.jsonl"


class FillReport:
    """Mutable per-fill outcome record. JSON-friendly by design."""

    def __init__(self, template_path="", output_path=""):
        self.template = os.path.basename(template_path)
        self.output = output_path
        self.timestamp = datetime.now(timezone.utc).isoformat()
        self.filled = 0                  # assignments applied to a text frame
        self.images_filled = 0
        self.skipped = 0                 # shapes explicitly marked skip
        self.out_of_range = 0            # slide/shape indices past the deck
        self.missing_metrics = []        # metric names not in this export
        self.unmatched_placeholders = [] # (metric, replace_text) never found
        self.missing_images = []         # image paths that didn't resolve
        self.failed_queries = []         # query-backed metrics that errored
        self.errors = []                 # unexpected per-shape exceptions

    # ── Recording ─────────────────────────────────────────────────────────

    def note_filled(self):
        self.filled += 1

    def note_image(self):
        self.images_filled += 1

    def note_skip(self):
        self.skipped += 1

    def note_out_of_range(self):
        self.out_of_range += 1

    def note_missing_metric(self, name):
        if name and name not in self.missing_metrics:
            self.missing_metrics.append(name)

    def note_unmatched_placeholder(self, metric, replace_text):
        self.unmatched_placeholders.append(
            {"metric": metric, "placeholder": replace_text})

    def note_missing_image(self, path):
        self.missing_images.append(str(path))

    def note_failed_query(self, metric):
        if metric and metric not in self.failed_queries:
            self.failed_queries.append(metric)

    def note_error(self, detail):
        self.errors.append(str(detail))

    # ── Reading ───────────────────────────────────────────────────────────

    @property
    def attempted(self):
        return (self.filled + len(self.missing_metrics)
                + len(self.unmatched_placeholders) + len(self.failed_queries))

    @property
    def ok(self):
        """True when nothing the user mapped was left unfilled."""
        return not (self.missing_metrics or self.unmatched_placeholders
                    or self.missing_images or self.failed_queries
                    or self.errors)

    def summary(self):
        """One human-readable paragraph for the post-fill dialog."""
        parts = [f"Filled {self.filled} text assignment(s)"]
        if self.images_filled:
            parts.append(f"{self.images_filled} image(s)")
        line = " and ".join(parts) + "."
        problems = []
        if self.missing_metrics:
            problems.append("Metrics not in this export: "
                            + ", ".join(self.missing_metrics))
        if self.unmatched_placeholders:
            ph = ", ".join(p["placeholder"] or p["metric"]
                           for p in self.unmatched_placeholders)
            problems.append(f"Placeholder text not found in template: {ph}")
        if self.missing_images:
            problems.append("Images not found: "
                            + ", ".join(os.path.basename(p)
                                        for p in self.missing_images))
        if self.failed_queries:
            problems.append("Queries that failed: "
                            + ", ".join(self.failed_queries))
        if self.errors:
            problems.append(f"{len(self.errors)} unexpected error(s) — see log")
        if self.skipped:
            problems.append(f"{self.skipped} shape(s) marked skip (by design)")
        if problems:
            line += "\n\n" + "\n".join(f"• {p}" for p in problems)
        return line

    def to_dict(self):
        return {
            "timestamp": self.timestamp,
            "template": self.template,
            "output": self.output,
            "ok": self.ok,
            "filled": self.filled,
            "images_filled": self.images_filled,
            "skipped": self.skipped,
            "out_of_range": self.out_of_range,
            "missing_metrics": list(self.missing_metrics),
            "unmatched_placeholders": list(self.unmatched_placeholders),
            "missing_images": list(self.missing_images),
            "failed_queries": list(self.failed_queries),
            "errors": list(self.errors),
        }


def append_fill_history(report):
    """Append one JSON line per fill to workspace/logs/fill_history.jsonl.

    Best-effort telemetry: any failure here is logged and swallowed — a
    broken history file must never break report generation.
    """
    try:
        from config.paths import LOGS_DIR
        os.makedirs(LOGS_DIR, exist_ok=True)
        path = os.path.join(LOGS_DIR, FILL_HISTORY_NAME)
        with open(path, "a", encoding="utf-8") as f:
            f.write(json.dumps(report.to_dict()) + "\n")
        return path
    except Exception:
        logger.warning("Could not append fill history", exc_info=True)
        return None
