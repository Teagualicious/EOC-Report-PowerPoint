"""Shared fixtures and import indirection for the test suite.

The suite is written to pass both before and after the module split:
each ``import_*`` helper tries the post-refactor module path first and
falls back to the pre-refactor location, so the same tests characterize
the monolith and then guard the refactor.
"""

import os
import sys
import textwrap

import pytest

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
APP_DIR = os.path.join(PROJECT_ROOT, "app")
if APP_DIR not in sys.path:
    sys.path.insert(0, APP_DIR)


# ── Import indirection ────────────────────────────────────────────────────────

def import_data_pipeline():
    """Module holding apply_platform_config / filter_data_by_campaigns / etc."""
    try:
        from engine import data_pipeline
        return data_pipeline
    except ImportError:
        import ingestion_engine
        return ingestion_engine


def import_excel_writer():
    try:
        from engine import excel_writer
        return excel_writer
    except ImportError:
        import excel_writer
        return excel_writer


def import_pptx_formats():
    """Module holding _format_value / _format_date / _detect_date_format."""
    try:
        from engine import pptx_formats
        return pptx_formats
    except ImportError:
        import pptx_mapper
        return pptx_mapper


def import_query_resolver():
    try:
        from engine import query_resolver
        return query_resolver
    except ImportError:
        import query_resolver
        return query_resolver


def import_kpi():
    """Post-refactor KPI module; None while the logic still lives in the UI."""
    try:
        from engine import kpi
        return kpi
    except ImportError:
        return None


def import_metrics_catalog():
    """Module holding get_available_metrics."""
    try:
        from engine import metrics_catalog
        return metrics_catalog
    except ImportError:
        import pptx_mapper
        return pptx_mapper


# ── Dictionary cache isolation ────────────────────────────────────────────────

@pytest.fixture(autouse=True)
def _reset_dictionary_cache():
    """The metric dictionary is cached module-globally; reset around each test."""
    from parsers import dictionary
    dictionary._cached_dict = None
    yield
    dictionary._cached_dict = None


@pytest.fixture(autouse=True)
def _isolate_fingerprints(tmp_path, monkeypatch):
    """HTMLParser persists structure fingerprints; keep them out of mappings/."""
    from parsers import html_parser
    monkeypatch.setattr(html_parser, "FINGERPRINT_DIR", str(tmp_path / "fp"))


# ── Synthetic sample files ────────────────────────────────────────────────────

@pytest.fixture
def sample_csv(tmp_path):
    """Campaign-level CSV in AudienceTrak style."""
    path = tmp_path / "audiencetrak_sample.csv"
    path.write_text(textwrap.dedent("""\
        Campaign,Impressions,Clicks,100% Completions,Spend
        Acme Appliance Co - Streaming TV,"1,204,556",3421,"1,100,203","$8,400.50"
        Acme Appliance Co - Display,88450,1205,0,"$1,200.00"
        """), encoding="utf-8")
    return str(path)


@pytest.fixture
def sample_xlsx(tmp_path):
    """Architect-style export: breakdown sheets only, no campaign summary."""
    from openpyxl import Workbook

    wb = Workbook()

    ws = wb.active
    ws.title = "Data by Devices"
    ws.append(["Order", "Device", "Impressions", "Clicks", "100% Complete"])
    ws.append(["Campaign A", "Roku", 20000, 150, 18000])
    ws.append(["Campaign A", "Samsung TV", 15000, 90, 13500])
    ws.append(["Campaign A", "Mobile", 5000, 60, 4100])
    ws.append(["Campaign B", "Roku", 30000, 210, 27000])
    ws.append(["Campaign B", "Mobile", 10000, 80, 8800])

    ws2 = wb.create_sheet("Data by Geo")
    ws2.append(["Order", "Zip Code", "Impressions"])
    ws2.append(["Campaign A", "33607", 22000])
    ws2.append(["Campaign A", "33609", 18100])
    ws2.append(["Campaign B", "33607", 39900])

    path = tmp_path / "architect_sample.xlsx"
    wb.save(str(path))
    return str(path)


@pytest.fixture
def sample_html(tmp_path):
    """Search-platform style HTML report mirroring a typical input report structure."""
    path = tmp_path / "report_sample.html"
    path.write_text(textwrap.dedent("""\
        <!DOCTYPE html>
        <html>
        <head><title>New Vehicle Search Campaign</title></head>
        <body>
        <h1>New Vehicle Search Campaign</h1>
        <p>Report Period: May 1, 2026 - May 31, 2026</p>
        <table>
            <tr><th>Metric</th><th>Value</th></tr>
            <tr><td>Impressions</td><td>52,180</td></tr>
            <tr><td>Clicks</td><td>2,103</td></tr>
            <tr><td>CTR</td><td>4.03%</td></tr>
            <tr><td>Cost</td><td>$3,678.42</td></tr>
        </table>
        </body>
        </html>
        """), encoding="utf-8")
    return str(path)


# ── Parsed-data fixtures (post-parse pipeline shapes) ────────────────────────

@pytest.fixture
def client_data():
    """Two-source client data exercising max-across-sources KPI logic.

    device breakdown sums: Impressions 50,000 / geo breakdown: 60,000
    → the KPI total for Impressions must be 60,000 (max), not the sum.
    """
    return [
        {
            "source_file": "architect.xlsx",
            "client_name": "Acme Appliance Co",
            "campaign_name": "Campaign A",
            "campaign_type": "",
            "source_platform": "Architect",
            "start_date": "2026-05-01",
            "end_date": "2026-05-31",
            "campaign_metrics": {},
            "mapped_metrics": {},
            "detected_tables": [],
            "level_data": [
                {"metric_level": "device:Roku", "metric_name": "Impressions",
                 "metric_value": 30000, "_campaign": "Campaign A"},
                {"metric_level": "device:Mobile", "metric_name": "Impressions",
                 "metric_value": 20000, "_campaign": "Campaign A"},
                {"metric_level": "zip:33607", "metric_name": "Impressions",
                 "metric_value": 35000, "_campaign": "Campaign A"},
                {"metric_level": "zip:33609", "metric_name": "Impressions",
                 "metric_value": 25000, "_campaign": "Campaign A"},
                {"metric_level": "device:Roku", "metric_name": "100% Completions",
                 "metric_value": 27000, "_campaign": "Campaign A"},
            ],
        },
        {
            "source_file": "audiencetrak.csv",
            "client_name": "Acme Appliance Co",
            "campaign_name": "Campaign B",
            "campaign_type": "",
            "source_platform": "AudienceTrak",
            "start_date": "2026-05-01",
            "end_date": "2026-05-31",
            "campaign_metrics": {
                "Campaign B|Impressions": {
                    "value": 88450, "universal_name": "Impressions",
                    "campaign_name": "Campaign B"},
                "Campaign B|Clicks": {
                    "value": 1205, "universal_name": "Clicks",
                    "campaign_name": "Campaign B"},
            },
            "mapped_metrics": {},
            "detected_tables": [],
            "level_data": [],
        },
    ]
