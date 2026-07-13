"""
HTML Parser with structure fingerprinting and table detection.
"""

from bs4 import BeautifulSoup
import logging
import re
import os
import json

from config.paths import MAPPINGS_DIR
from config.naming import storage_key
from config.settings import _atomic_json_write
from engine.errors import ParserError

logger = logging.getLogger(__name__)

FINGERPRINT_DIR = MAPPINGS_DIR


class HTMLParser:
    """Parses an HTML report using structure fingerprints + table detection."""

    def __init__(self, filepath, platform_name="default"):
        self.filepath = filepath
        self.platform_name = platform_name

    def parse(self):
        """Parse the file. Raises ParserError if it cannot be read."""
        if not os.path.exists(self.filepath):
            raise ParserError(f"File not found: {self.filepath}",
                              user_message="The file no longer exists at "
                                           f"{self.filepath}")
        try:
            content = None
            for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
                try:
                    with open(self.filepath, "r", encoding=encoding, errors="strict") as f:
                        content = f.read()
                    break
                except UnicodeDecodeError:
                    continue
            if content is None:
                with open(self.filepath, "r", encoding="utf-8", errors="replace") as f:
                    content = f.read()
            soup = BeautifulSoup(content, "html.parser")
        except OSError as e:
            raise ParserError(f"Cannot read {self.filepath}: {e}",
                              user_message="The file could not be opened — is it "
                                           "locked by another program?") from e

        fingerprint = self._load_fingerprint()
        if fingerprint:
            metrics = self._extract_with_fingerprint(soup, fingerprint)
            if len(metrics) < len(fingerprint):
                full = self._full_scan(soup)
                for k, v in full.items():
                    if k not in metrics: metrics[k] = v
                self._save_fingerprint(self._build_fingerprint(soup, metrics))
        else:
            metrics = self._full_scan(soup)
            self._save_fingerprint(self._build_fingerprint(soup, metrics))

        detected_tables = self._detect_tables(soup)

        # Classify detected tables into the unified schema, exactly like the
        # CSV/Excel parsers do. Previously HTML relied entirely on a platform
        # config rebuild — without one (or with mismatched sheet names) the
        # returned dict had no level_data/campaign_metrics keys and HTML
        # files contributed zero rows (or crashed the pipeline).
        from parsers.dictionary import (classify_columns, extract_row_context,
                                        build_level_rows, build_campaign_rows)
        level_data = []
        campaign_metrics = {}
        for tbl in detected_tables:
            headers = tbl.get("headers", [])
            classification = classify_columns(headers)
            found_context = classification["context"]
            found_levels = classification["levels"]
            found_metrics = classification["metrics"]
            tbl["found_levels"] = [(cn, ld["prefix"]) for cn, ld in found_levels]
            tbl["found_metrics"] = found_metrics
            tbl["found_context"] = found_context
            for row in tbl.get("rows", []):
                ctx = extract_row_context(row, found_context)
                if found_levels:
                    level_data.extend(
                        build_level_rows(row, found_levels, found_metrics, ctx))
                else:
                    campaign_metrics.update(
                        build_campaign_rows(row, found_metrics, ctx))

        return {
            "source_file": os.path.basename(self.filepath),
            "campaign_name": self._extract_campaign_name(soup),
            "date_range": self._extract_date_range(soup),
            "metrics": metrics,
            "detected_tables": detected_tables,
            "level_data": level_data,
            "campaign_metrics": campaign_metrics,
        }

    # ── Table detection ──────────────────────────────────────────────────

    def _detect_tables(self, soup):
        """Find all HTML tables and extract headers + rows as dicts."""
        detected = []
        for table in soup.find_all("table"):
            headers = []
            thead = table.find("thead")
            if thead:
                for th in thead.find_all(["th", "td"]):
                    headers.append(th.get_text(strip=True))
            else:
                first_row = table.find("tr")
                if first_row:
                    for cell in first_row.find_all(["th", "td"]):
                        headers.append(cell.get_text(strip=True))

            if not headers or len(headers) < 2:
                continue

            rows = []
            tbody = table.find("tbody") or table
            for tr in tbody.find_all("tr"):
                cells = tr.find_all(["td", "th"])
                if len(cells) < 2:
                    continue
                values = [c.get_text(strip=True) for c in cells]
                if values == headers:
                    continue  # skip header row if repeated in tbody

                row_dict = {}
                for i, h in enumerate(headers):
                    if i < len(values):
                        row_dict[h] = self._clean_value(values[i])
                if any(v != "" for v in row_dict.values()):
                    rows.append(row_dict)

            if rows:
                detected.append({
                    "headers": headers,
                    "rows": rows,
                    "row_count": len(rows),
                })
        return detected

    # ── Fingerprint ──────────────────────────────────────────────────────

    def _fingerprint_path(self):
        safe = storage_key(self.platform_name)
        return os.path.join(FINGERPRINT_DIR, f"fingerprint_{safe}.json")

    def _load_fingerprint(self):
        path = self._fingerprint_path()
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    return json.load(f)
            except (OSError, json.JSONDecodeError):
                logger.warning("Corrupt fingerprint file ignored: %s", path,
                               exc_info=True)
        return None

    def _save_fingerprint(self, fp):
        os.makedirs(FINGERPRINT_DIR, exist_ok=True)
        _atomic_json_write(self._fingerprint_path(), fp)

    def _get_css_path(self, element):
        parts = []
        for parent in element.parents:
            if parent.name is None: break
            idx = 0
            for sibling in parent.parent.children if parent.parent else []:
                if hasattr(sibling, "name") and sibling.name == parent.name:
                    idx += 1
                    if sibling is parent: break
            classes = ".".join(parent.get("class", []))
            tag = parent.name
            if classes: tag += f".{classes}"
            parts.append(f"{tag}:nth-of-type({idx})")
        parts.reverse()
        return " > ".join(parts[-4:])

    def _build_fingerprint(self, soup, metrics):
        fp = {}
        for table in soup.find_all("table"):
            for row in table.find_all("tr"):
                cells = row.find_all(["td", "th"])
                if len(cells) >= 2:
                    label = cells[0].get_text(strip=True)
                    if label in metrics:
                        fp[label] = {"strategy": "table",
                                     "path": self._get_css_path(table),
                                     "label_text": label}
        for elem in soup.find_all(class_=re.compile(r"metric|stat|kpi", re.I)):
            label_elem = elem.find(class_=re.compile(r"label|name|title|header", re.I))
            if label_elem:
                label = label_elem.get_text(strip=True)
                if label in metrics and label not in fp:
                    fp[label] = {"strategy": "class_pattern",
                                 "path": self._get_css_path(elem),
                                 "label_text": label}
        return fp

    def _extract_with_fingerprint(self, soup, fingerprint):
        """Resolve fingerprint labels with one document scan.

        The previous implementation rescanned every table and KPI element once
        per metric (O(metrics × elements)); large HTML exports became noticeably
        slower after a fingerprint was learned.
        """
        table_values = {}
        for table in soup.find_all("table"):
            for row in table.find_all("tr"):
                cells = row.find_all(["td", "th"])
                if len(cells) >= 2:
                    label = cells[0].get_text(strip=True)
                    table_values.setdefault(label,
                        self._clean_value(cells[1].get_text(strip=True)))

        class_values = {}
        for elem in soup.find_all(class_=re.compile(r"metric|stat|kpi", re.I)):
            label_elem = elem.find(class_=re.compile(r"label|name|title|header", re.I))
            value_elem = elem.find(class_=re.compile(r"value|number|data|amount", re.I))
            if label_elem and value_elem:
                class_values.setdefault(label_elem.get_text(strip=True),
                    self._clean_value(value_elem.get_text(strip=True)))

        metrics = {}
        for label, info in fingerprint.items():
            target = info.get("label_text", label)
            strategy = info.get("strategy")
            if strategy == "table" and target in table_values:
                metrics[label] = table_values[target]
            elif strategy == "class_pattern" and target in class_values:
                metrics[label] = class_values[target]
        return metrics

    # ── Full scan ────────────────────────────────────────────────────────

    def _extract_campaign_name(self, soup):
        title = soup.find("title")
        if title: return title.get_text(strip=True)
        h1 = soup.find("h1")
        if h1: return h1.get_text(strip=True)
        return "Unknown Campaign"

    def _extract_date_range(self, soup):
        text = soup.get_text()
        for pattern in [
            r"(\w+ \d{1,2},?\s*\d{4})\s*(?:-|\u2013|\u2014|to)\s*(\w+ \d{1,2},?\s*\d{4})",
            r"(\d{1,2}/\d{1,2}/\d{2,4})\s*(?:-|\u2013|\u2014|to)\s*(\d{1,2}/\d{1,2}/\d{2,4})",
            r"(\d{4}-\d{2}-\d{2})\s*(?:-|\u2013|\u2014|to)\s*(\d{4}-\d{2}-\d{2})",
        ]:
            match = re.search(pattern, text)
            if match: return f"{match.group(1)} - {match.group(2)}"
        return "Unknown Date Range"

    def _full_scan(self, soup):
        metrics = {}
        for table in soup.find_all("table"):
            for row in table.find_all("tr"):
                cells = row.find_all(["td", "th"])
                if len(cells) >= 2:
                    label = cells[0].get_text(strip=True)
                    value = cells[1].get_text(strip=True)
                    if label and value and self._looks_like_metric(label, value):
                        metrics[label] = self._clean_value(value)

        class_pattern = re.compile(
            r"metric|stat|kpi|value|measure|score|count|total|summary", re.I)
        for elem in soup.find_all(class_=class_pattern):
            le = elem.find(class_=re.compile(r"label|name|title|header", re.I))
            ve = elem.find(class_=re.compile(r"value|number|data|amount", re.I))
            if le and ve:
                label = le.get_text(strip=True)
                value = ve.get_text(strip=True)
                if label and value:
                    metrics[label] = self._clean_value(value)

        for dl in soup.find_all("dl"):
            for dt, dd in zip(dl.find_all("dt"), dl.find_all("dd")):
                label = dt.get_text(strip=True); value = dd.get_text(strip=True)
                if label and value: metrics[label] = self._clean_value(value)

        for elem in soup.find_all(["div", "span", "p"]):
            text = elem.get_text(strip=True)
            match = re.match(r"^([A-Za-z\s]+)[:]\s*([0-9,.%$]+)$", text)
            if match: metrics[match.group(1).strip()] = self._clean_value(match.group(2))

        return metrics

    def _looks_like_metric(self, label, value):
        return bool(re.search(r"[0-9]", value) or "%" in value or "$" in value)

    def _clean_value(self, value_str):
        cleaned = value_str.strip()
        numeric = cleaned.replace("$", "").replace(",", "").replace("%", "").strip()
        try:
            return float(numeric) if "." in numeric else int(numeric)
        except ValueError:
            return cleaned
