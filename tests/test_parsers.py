"""Tests for the CSV, Excel, and HTML parsers — happy paths and bad input."""

import pytest

from parsers.csv_parser import CSVParser
from parsers.excel_parser import ExcelParser
from parsers.html_parser import HTMLParser


class TestCSVParser:
    def test_campaign_level_parse(self, sample_csv):
        result = CSVParser(sample_csv).parse()
        cm = result["campaign_metrics"]
        assert cm["Acme Appliance Co - Streaming TV|Impressions"]["value"] == 1204556
        assert cm["Acme Appliance Co - Streaming TV|Cost"]["value"] == 8400.50
        # "100% Completions" resolves through the shipped dictionary
        assert cm["Acme Appliance Co - Streaming TV|100% Completions"]["value"] == 1100203
        assert result["campaign_name"] == "Acme Appliance Co - Streaming TV"
        assert result["level_data"] == []

    def test_number_cleaning(self, sample_csv):
        result = CSVParser(sample_csv).parse()
        val = result["campaign_metrics"]["Acme Appliance Co - Display|Impressions"]["value"]
        assert isinstance(val, int) and val == 88450

    def test_empty_file(self, tmp_path):
        path = tmp_path / "empty.csv"
        path.write_text("", encoding="utf-8")
        result = CSVParser(str(path)).parse()
        assert result["campaign_metrics"] == {}
        assert result["level_data"] == []

    def test_header_only_file(self, tmp_path):
        path = tmp_path / "headers.csv"
        path.write_text("Campaign,Impressions\n", encoding="utf-8")
        result = CSVParser(str(path)).parse()
        assert result["campaign_metrics"] == {}

    def test_utf8_bom_does_not_corrupt_first_header(self, tmp_path):
        path = tmp_path / "bom.csv"
        path.write_text(
            "Campaign,Impressions\nBOM Campaign,1000\n",
            encoding="utf-8-sig")
        result = CSVParser(str(path)).parse()
        assert result["campaign_name"] == "BOM Campaign"
        assert "BOM Campaign|Impressions" in result["campaign_metrics"]

    def test_latin1_encoding(self, tmp_path):
        path = tmp_path / "latin.csv"
        path.write_bytes("Campaign,Impressions\nCafé Motors,1000\n".encode("latin-1"))
        result = CSVParser(str(path)).parse()
        assert len(result["campaign_metrics"]) == 1

    def test_table_only_mode_skips_duplicate_unified_rows(self, sample_csv):
        result = CSVParser(sample_csv).parse(build_outputs=False)
        assert result["detected_tables"]
        assert result["campaign_metrics"] == {}
        assert result["level_data"] == []

    def test_level_breakdown_csv(self, tmp_path):
        path = tmp_path / "devices.csv"
        path.write_text(
            "Campaign,Device,Impressions\nCamp A,Roku,5000\nCamp A,Mobile,3000\n",
            encoding="utf-8")
        result = CSVParser(str(path)).parse()
        assert len(result["level_data"]) == 2
        assert result["level_data"][0]["metric_level"] == "device:Roku"


class TestExcelParser:
    def test_multi_sheet_parse(self, sample_xlsx):
        result = ExcelParser(sample_xlsx).parse()
        levels = {ld["metric_level"] for ld in result["level_data"]}
        assert "device:Roku" in levels
        assert "zip:33607" in levels
        campaigns = {ld["_campaign"] for ld in result["level_data"]}
        assert campaigns == {"Campaign A", "Campaign B"}

    def test_table_only_mode_skips_duplicate_unified_rows(self, sample_xlsx):
        result = ExcelParser(sample_xlsx).parse(build_outputs=False)
        assert result["detected_tables"]
        assert result["campaign_metrics"] == {}
        assert result["level_data"] == []

    def test_values_parsed_numeric(self, sample_xlsx):
        result = ExcelParser(sample_xlsx).parse()
        roku_a = [ld for ld in result["level_data"]
                  if ld["metric_level"] == "device:Roku"
                  and ld["_campaign"] == "Campaign A"
                  and ld["metric_name"] == "Impressions"]
        assert roku_a[0]["metric_value"] == 20000

    def test_two_tables_detected(self, sample_xlsx):
        result = ExcelParser(sample_xlsx).parse()
        names = {t["sheet_name"] for t in result["detected_tables"]}
        assert names == {"Data by Devices", "Data by Geo"}

    def test_empty_workbook(self, tmp_path):
        from openpyxl import Workbook
        path = tmp_path / "empty.xlsx"
        Workbook().save(str(path))
        result = ExcelParser(str(path)).parse()
        assert result["detected_tables"] == []
        assert result["level_data"] == []

    def test_header_detection_with_preamble(self, tmp_path):
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.append(["Monthly Report"])
        ws.append([])
        ws.append(["Order", "Device", "Impressions"])
        ws.append(["Camp X", "Roku", 1234])
        path = tmp_path / "preamble.xlsx"
        wb.save(str(path))
        result = ExcelParser(str(path)).parse()
        assert result["detected_tables"][0]["headers"] == \
            ["Order", "Device", "Impressions"]


class TestHTMLParser:
    def test_metric_table_parse(self, sample_html):
        result = HTMLParser(sample_html, "test_platform").parse()
        assert result["metrics"]["Impressions"] == 52180
        assert result["metrics"]["Cost"] == 3678.42
        assert result["campaign_name"] == "New Vehicle Search Campaign"

    def test_date_range_extracted(self, sample_html):
        result = HTMLParser(sample_html, "test_platform").parse()
        assert "May 1, 2026" in result["date_range"]

    def test_date_range_accepts_word_separator(self, tmp_path):
        path = tmp_path / "date_to.html"
        path.write_text(
            "<html><body>May 1, 2026 to May 31, 2026</body></html>",
            encoding="utf-8")
        result = HTMLParser(str(path), "test_platform").parse()
        assert result["date_range"] == "May 1, 2026 - May 31, 2026"

    def test_date_range_rejects_single_separator_letter(self, tmp_path):
        path = tmp_path / "bad_date.html"
        path.write_text(
            "<html><body>May 1, 2026 t May 31, 2026</body></html>",
            encoding="utf-8")
        result = HTMLParser(str(path), "test_platform").parse()
        assert result["date_range"] == "Unknown Date Range"

    def test_detected_tables(self, sample_html):
        result = HTMLParser(sample_html, "test_platform").parse()
        assert result["detected_tables"][0]["row_count"] == 4

    def test_empty_html(self, tmp_path):
        path = tmp_path / "empty.html"
        path.write_text("<html><body></body></html>", encoding="utf-8")
        result = HTMLParser(str(path), "test_platform").parse()
        assert result["metrics"] == {}
        assert result["campaign_name"] == "Unknown Campaign"

    def test_real_input_files(self):
        """Any optional HTML samples in root input/ must keep parsing."""
        import os
        from config import paths
        base = paths.INPUT_DIR
        samples = [name for name in os.listdir(base)
                   if name.lower().endswith((".html", ".htm"))]
        for name in samples:
            result = HTMLParser(os.path.join(base, name), "smoke").parse()
            assert result["metrics"], f"{name} produced no metrics"
