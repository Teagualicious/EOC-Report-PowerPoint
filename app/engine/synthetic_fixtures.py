"""Deterministic, fake campaign-export fixtures for Stage 1 and later stages.

The factory is deliberately production-independent: it creates only obviously
fake campaign identifiers and values, never reads a client file, and uses no
clock or random source.  Tests and local demonstrations can extend the rows
without creating a second fixture convention.

Public API:

* :func:`synthetic_campaign_rows` returns deterministic dictionaries;
* :func:`write_synthetic_csv`, :func:`write_synthetic_xlsx`, and
  :func:`write_synthetic_html` write the same shape in each supported format;
* :func:`write_synthetic_export` dispatches by the destination suffix.
"""

from __future__ import annotations

import csv
import html
import os
from pathlib import Path
from typing import Any, Iterable, Mapping


SYNTHETIC_HEADERS = ("Campaign", "Impressions", "Clicks", "Cost")


def synthetic_campaign_rows(count: int = 2) -> list[dict[str, Any]]:
    """Return ``count`` deterministic fake campaign rows."""
    if count < 0:
        raise ValueError("count must be non-negative")
    rows = []
    for index in range(count):
        rows.append({
            "Campaign": f"FAKE-{index + 1:03d}",
            "Impressions": 10000 + index * 1375,
            "Clicks": 250 + index * 31,
            "Cost": round(250.0 + index * 17.5, 2),
        })
    return rows


def _destination(path: str | os.PathLike[str]) -> Path:
    destination = Path(path)
    destination.parent.mkdir(parents=True, exist_ok=True)
    return destination


def write_synthetic_csv(path: str | os.PathLike[str], count: int = 2,
                        rows: Iterable[Mapping[str, Any]] | None = None) -> str:
    """Write a deterministic fake campaign CSV and return its path."""
    destination = _destination(path)
    data = list(rows) if rows is not None else synthetic_campaign_rows(count)
    with destination.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(SYNTHETIC_HEADERS))
        writer.writeheader()
        writer.writerows({header: row.get(header, "") for header in SYNTHETIC_HEADERS}
                         for row in data)
    return str(destination)


def write_synthetic_xlsx(path: str | os.PathLike[str], count: int = 2,
                         rows: Iterable[Mapping[str, Any]] | None = None) -> str:
    """Write a deterministic fake campaign workbook and return its path."""
    from openpyxl import Workbook

    destination = _destination(path)
    data = list(rows) if rows is not None else synthetic_campaign_rows(count)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Campaign Summary"
    sheet.append(list(SYNTHETIC_HEADERS))
    for row in data:
        sheet.append([row.get(header, "") for header in SYNTHETIC_HEADERS])
    workbook.save(destination)
    workbook.close()
    return str(destination)


def write_synthetic_html(path: str | os.PathLike[str], count: int = 2,
                         rows: Iterable[Mapping[str, Any]] | None = None) -> str:
    """Write a deterministic fake campaign HTML table and return its path."""
    destination = _destination(path)
    data = list(rows) if rows is not None else synthetic_campaign_rows(count)
    header_html = "".join(f"<th>{html.escape(header)}</th>"
                          for header in SYNTHETIC_HEADERS)
    row_html = []
    for row in data:
        cells = "".join(
            f"<td>{html.escape(str(row.get(header, '')))}</td>"
            for header in SYNTHETIC_HEADERS
        )
        row_html.append(f"<tr>{cells}</tr>")
    content = (
        "<!DOCTYPE html><html><head><title>Fake Campaign Export</title></head>"
        f"<body><table><thead><tr>{header_html}</tr></thead><tbody>"
        f"{''.join(row_html)}</tbody></table></body></html>"
    )
    destination.write_text(content, encoding="utf-8")
    return str(destination)


def write_synthetic_export(path: str | os.PathLike[str], count: int = 2,
                           rows: Iterable[Mapping[str, Any]] | None = None) -> str:
    """Write the deterministic fixture format selected by *path* suffix."""
    suffix = Path(path).suffix.lower()
    if suffix == ".csv":
        return write_synthetic_csv(path, count=count, rows=rows)
    if suffix in {".xlsx", ".xlsm"}:
        return write_synthetic_xlsx(path, count=count, rows=rows)
    if suffix in {".html", ".htm"}:
        return write_synthetic_html(path, count=count, rows=rows)
    raise ValueError("Synthetic exports must use .csv, .xlsx, .xlsm, .html, or .htm")
