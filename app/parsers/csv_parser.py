"""
CSV Parser with shared dictionary integration.
Uses the same alias matching, context columns, and level detection as the Excel parser.
"""

import csv
import logging
import os

from engine.errors import ParserError
from parsers.dictionary import (
    classify_columns, extract_row_context,
    build_level_rows, build_campaign_rows
)

logger = logging.getLogger(__name__)


class CSVParser:
    """Parses a CSV export into the unified parsed-data dict."""

    def __init__(self, filepath):
        self.filepath = filepath

    def parse(self, build_outputs=True):
        """Parse the file. Raises ParserError if it cannot be read.

        ``build_outputs=False`` keeps only the detected source table. The main
        workflow uses that mode because its saved platform configuration will
        rebuild the unified rows immediately afterward, avoiding duplicate work.
        """
        if not os.path.exists(self.filepath):
            raise ParserError(f"File not found: {self.filepath}",
                              user_message="The file no longer exists at "
                                           f"{self.filepath}")
        # Detect encoding from a bounded sample, then stream rows from disk.
        # utf-8-sig must be attempted before utf-8 so a BOM does not become
        # part of the first header and silently break dictionary matching.
        encoding = None
        sample = ""
        try:
            for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
                try:
                    with open(self.filepath, "r", encoding=enc, errors="strict",
                              newline="") as f:
                        sample = f.read(65536)
                    encoding = enc
                    break
                except UnicodeDecodeError:
                    continue
        except OSError as e:
            raise ParserError(f"Cannot read {self.filepath}: {e}",
                              user_message="The file could not be opened — is it "
                                           "locked by another program?") from e

        if encoding is None:
            encoding = "utf-8"
            logger.warning("No encoding matched %s — falling back to utf-8 with "
                           "errors ignored", self.filepath)

        try:
            delimiter = csv.Sniffer().sniff(sample[:4096]).delimiter
        except csv.Error:
            delimiter = ","

        rows = []
        try:
            with open(self.filepath, "r", encoding=encoding,
                      errors="ignore" if encoding == "utf-8" and not sample else "strict",
                      newline="") as f:
                reader = csv.DictReader(f, delimiter=delimiter, strict=True)
                fieldnames = reader.fieldnames or []
                normalized_headers = [
                    " ".join(str(field or "").strip().split()).casefold()
                    for field in fieldnames
                ]
                if any(not header for header in normalized_headers):
                    raise ParserError(
                        "CSV header row contains an empty column name",
                        user_message=(
                            "The CSV header row contains an empty column name. "
                            "Name every column and export it again."
                        ),
                        code="HEADER_REQUIRED",
                    )
                duplicates = sorted({
                    header for header in normalized_headers
                    if header and normalized_headers.count(header) > 1
                })
                if duplicates:
                    raise ParserError(
                        f"Duplicate CSV headers: {', '.join(duplicates)}",
                        user_message=(
                            "The CSV contains duplicate column names: "
                            f"{', '.join(duplicates)}. Rename them and export again."
                        ),
                        code="DUPLICATE_HEADERS",
                    )
                for row_number, row in enumerate(reader, start=2):
                    if None in row:
                        raise ParserError(
                            f"CSV row {row_number} has more values than its header",
                            user_message=(
                                f"CSV row {row_number} has more values than its "
                                "header. Repair the export and try again."
                            ),
                            code="MALFORMED_CSV",
                        )
                    cleaned = {}
                    for key, value in row.items():
                        if key is not None:
                            cleaned[key.strip()] = value.strip() if value else ""
                    rows.append(cleaned)
        except csv.Error as e:
            raise ParserError(
                f"Malformed CSV {self.filepath}: {e}",
                user_message=(
                    "The CSV export is malformed. Re-export it with a standard "
                    "comma- or tab-delimited table."
                ),
                code="MALFORMED_CSV",
            ) from e
        except OSError as e:
            raise ParserError(f"Cannot read {self.filepath}: {e}",
                              user_message="The file could not be opened — is it "
                                           "locked by another program?") from e

        if not rows:
            return {
                "source_file": os.path.basename(self.filepath),
                "campaign_name": os.path.splitext(os.path.basename(self.filepath))[0],
                "date_range": "See file",
                "metrics": {},
                "detected_tables": [],
                "level_data": [],
                "campaign_metrics": {},
            }

        headers = list(rows[0].keys())

        # Clean values
        for row in rows:
            for k in row:
                row[k] = self._clean_value(row[k])

        # Classify columns using shared dictionary
        classification = classify_columns(headers)
        found_context = classification["context"]
        found_levels = classification["levels"]
        found_metrics = classification["metrics"]

        # Build outputs
        level_data = []
        campaign_metrics = {}
        detected_tables = []

        if build_outputs:
            if found_levels:
                # Breakdown data — build level rows
                for row in rows:
                    ctx = extract_row_context(row, found_context)
                    level_rows = build_level_rows(row, found_levels, found_metrics, ctx)
                    level_data.extend(level_rows)
            else:
                # Campaign-level data
                for row in rows:
                    ctx = extract_row_context(row, found_context)
                    cm = build_campaign_rows(row, found_metrics, ctx)
                    campaign_metrics.update(cm)

        # Build flat metrics for mapping UI
        flat_metrics = {}
        for mdata in campaign_metrics.values():
            flat_metrics[mdata["universal_name"]] = mdata["value"]

        # Build detected table for UI
        detected_tables.append({
            "headers": headers,
            "rows": rows,
            "row_count": len(rows),
            "sheet_name": os.path.splitext(os.path.basename(self.filepath))[0],
            "found_levels": [(cn, ld["prefix"]) for cn, ld in found_levels],
            "found_metrics": found_metrics,
            "found_context": found_context,
        })

        # Extract campaign name
        campaign_name = ""
        for mdata in campaign_metrics.values():
            if mdata.get("campaign_name"):
                campaign_name = mdata["campaign_name"]
                break
        if not campaign_name:
            campaign_name = os.path.splitext(os.path.basename(self.filepath))[0]

        return {
            "source_file": os.path.basename(self.filepath),
            "campaign_name": campaign_name,
            "date_range": "See file",
            "metrics": flat_metrics,
            "detected_tables": detected_tables,
            "level_data": level_data,
            "campaign_metrics": campaign_metrics,
        }

    def _clean_value(self, value_str):
        if not isinstance(value_str, str):
            return value_str
        cleaned = value_str.strip()
        if not cleaned:
            return ""
        numeric = cleaned.replace("$", "").replace(",", "").replace("%", "").strip()
        try:
            if "." in numeric:
                return float(numeric)
            return int(numeric)
        except (ValueError, TypeError):
            return cleaned
