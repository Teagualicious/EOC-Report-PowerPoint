"""Tests for the data pipeline: file scanning, platform config, filtering."""

from conftest import import_data_pipeline

pipeline = import_data_pipeline()


class TestScanFileStructure:
    def test_scan_csv(self, sample_csv):
        sheets = pipeline.scan_file_structure(sample_csv)
        assert len(sheets) == 1
        assert "Impressions" in sheets[0]["headers"]
        assert sheets[0]["row_count"] == 2
        assert sheets[0]["classification"]["context"] == {"campaign_name": "Campaign"}

    def test_scan_xlsx(self, sample_xlsx):
        sheets = pipeline.scan_file_structure(sample_xlsx)
        names = {s["name"] for s in sheets}
        assert names == {"Data by Devices", "Data by Geo"}

    def test_scan_html(self, sample_html):
        sheets = pipeline.scan_file_structure(sample_html)
        assert any(s["name"] == "HTML Report" for s in sheets)

    def test_scan_unknown_extension(self, tmp_path):
        path = tmp_path / "file.txt"
        path.write_text("hello", encoding="utf-8")
        assert pipeline.scan_file_structure(str(path)) == []


def _device_sheet_config():
    return {
        "sheets": [{
            "sheet_name": "Data by Devices",
            "columns": [
                {"column": "Order", "role": "campaign_id", "selected": True},
                {"column": "Device", "role": "level: device", "selected": True},
                {"column": "Impressions", "role": "metric", "selected": True},
                {"column": "Clicks", "role": "metric", "selected": False},
            ],
        }]
    }


class TestApplyPlatformConfig:
    def test_builds_level_data(self, sample_xlsx):
        from parsers.excel_parser import ExcelParser
        parsed = ExcelParser(sample_xlsx).parse()
        pipeline.apply_platform_config(parsed, _device_sheet_config())
        levels = parsed["level_data"]
        assert all(ld["metric_level"].startswith("device:") for ld in levels)
        # Unselected column Clicks must not appear
        assert all(ld["metric_name"] != "Clicks" for ld in levels)
        roku_a = [ld for ld in levels if ld["metric_level"] == "device:Roku"
                  and ld["_campaign"] == "Campaign A"]
        assert roku_a[0]["metric_value"] == 20000

    def test_campaign_id_composition(self, tmp_path):
        parsed = {"detected_tables": [{
            "sheet_name": "Sheet1",
            "rows": [{"Advertiser": "Famous Tate", "Order": "May VLA",
                      "Impressions": 100}],
        }]}
        config = {"sheets": [{
            "sheet_name": "Sheet1",
            "columns": [
                {"column": "Advertiser", "role": "campaign_id", "selected": True},
                {"column": "Order", "role": "campaign_id", "selected": True},
                {"column": "Impressions", "role": "metric", "selected": True},
            ],
        }]}
        pipeline.apply_platform_config(parsed, config)
        key = list(parsed["campaign_metrics"].keys())[0]
        assert key == "Famous Tate - May VLA|Impressions"

    def test_total_rows_skipped(self):
        parsed = {"detected_tables": [{
            "sheet_name": "Sheet1",
            "rows": [
                {"Campaign": "Camp A", "Impressions": 100},
                {"Campaign": "Grand Total", "Impressions": 100},
            ],
        }]}
        config = {"sheets": [{
            "sheet_name": "Sheet1",
            "columns": [
                {"column": "Campaign", "role": "campaign_id", "selected": True},
                {"column": "Impressions", "role": "metric", "selected": True},
            ],
        }]}
        pipeline.apply_platform_config(parsed, config)
        assert list(parsed["campaign_metrics"]) == ["Camp A|Impressions"]

    def test_fuzzy_sheet_name_match(self):
        parsed = {"detected_tables": [{
            "sheet_name": "Data by Devices (2)",
            "rows": [{"Device": "Roku", "Impressions": 10}],
        }]}
        pipeline.apply_platform_config(parsed, _device_sheet_config())
        assert len(parsed["level_data"]) == 1

    def test_none_config_is_noop(self):
        parsed = {"detected_tables": []}
        pipeline.apply_platform_config(parsed, None)
        assert "campaign_metrics" not in parsed


    def test_unmatched_config_preserves_parser_output(self):
        original = {
            "source_file": "report.html",
            "campaign_metrics": {
                "Camp|Impressions": {"value": 10, "universal_name": "Impressions",
                                     "campaign_name": "Camp"}},
            "level_data": [],
            "detected_tables": [],
            "metrics": {},
        }
        config = {"sheets": [{"sheet_name": "Missing", "columns": []}]}
        pipeline.apply_platform_config(original, config)
        assert "Camp|Impressions" in original["campaign_metrics"]

    def test_metric_only_html_uses_campaign_name_fallback(self):
        parsed = {
            "source_file": "dashboard.html",
            "campaign_name": "Summer Campaign",
            "metrics": {"Impressions": 1000},
            "detected_tables": [],
            "campaign_metrics": {},
            "level_data": [],
        }
        config = {"sheets": [{
            "sheet_name": "HTML Report",
            "columns": [
                {"column": "Impressions", "role": "metric", "selected": True},
            ],
        }]}
        pipeline.apply_platform_config(parsed, config)
        assert parsed["campaign_metrics"][
            "Summer Campaign|Impressions"]["value"] == 1000

    def test_metric_alias_resolved(self):
        parsed = {"detected_tables": [{
            "sheet_name": "Sheet1",
            "rows": [{"Campaign": "Camp A", "Imps": 42}],
        }]}
        config = {"sheets": [{
            "sheet_name": "Sheet1",
            "columns": [
                {"column": "Campaign", "role": "campaign_id", "selected": True},
                {"column": "Imps", "role": "metric", "selected": True},
            ],
        }]}
        pipeline.apply_platform_config(parsed, config)
        # "Imps" resolves to universal name "Impressions"
        assert parsed["campaign_metrics"]["Camp A|Impressions"]["value"] == 42


class TestCampaignHelpers:

    def test_filter_keeps_only_selected(self, client_data):
        filtered = pipeline.filter_data_by_campaigns(client_data, ["Campaign A"])
        assert len(filtered) == 1
        assert all(ld["_campaign"] == "Campaign A"
                   for ld in filtered[0]["level_data"])

    def test_filter_empty_selection(self, client_data):
        assert pipeline.filter_data_by_campaigns(client_data, []) == []

    def test_filter_does_not_mutate_source(self, client_data):
        before = len(client_data[0]["level_data"])
        pipeline.filter_data_by_campaigns(client_data, ["Campaign A"])
        assert len(client_data[0]["level_data"]) == before
