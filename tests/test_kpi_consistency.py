"""Regression tests: KPI consistency between review screen and PowerPoint
totals, and HTML parser self-classification (v16.3 fixes)."""

from engine.kpi import compute_kpis
from engine.metrics_catalog import get_available_metrics


def _mixed_client_data():
    """CampA has a summary + two breakdowns (device is best); CampB has
    ONLY a creative breakdown. Ground truth total = 1005 + 500 = 1505."""
    return [{
        "campaign_metrics": {
            "a_imp": {"universal_name": "Impressions", "value": 1000,
                       "campaign_name": "CampA"},
        },
        "level_data": (
            [{"metric_name": "Impressions", "metric_value": v,
              "metric_level": f"network:N{i}", "_campaign": "CampA"}
             for i, v in enumerate([500, 490])] +
            [{"metric_name": "Impressions", "metric_value": v,
              "metric_level": f"device:D{i}", "_campaign": "CampA"}
             for i, v in enumerate([600, 405])] +
            [{"metric_name": "Impressions", "metric_value": v,
              "metric_level": f"creative:C{i}", "_campaign": "CampB"}
             for i, v in enumerate([300, 200])]
        ),
    }]


def test_review_kpi_uses_best_source_per_campaign():
    kpi_totals, details, _ = compute_kpis(_mixed_client_data())
    assert details["CampA"]["Impressions"] == 1005  # device beats summary
    assert details["CampB"]["Impressions"] == 500   # breakdown-only campaign
    assert kpi_totals["Impressions"] == 1505


def test_powerpoint_total_matches_review_total():
    """Regression: metrics_catalog used a GLOBAL best source, silently
    dropping campaigns whose best source differed (CampB vanished)."""
    metrics, _ = get_available_metrics(_mixed_client_data(), "T", "", "")
    assert metrics["Total Impressions"] == 1505


def test_powerpoint_totals_are_native_types():
    metrics, _ = get_available_metrics(_mixed_client_data(), "T", "", "")
    for k, v in metrics.items():
        if not k.startswith("__"):
            assert type(v) in (int, float, str), f"{k} leaked {type(v)}"


def test_html_parser_self_classifies(tmp_path):
    """Regression: HTML parser returned no level_data/campaign_metrics keys
    at all — HTML files crashed the pipeline or contributed zero rows."""
    html = ("<html><body><table>"
            "<tr><th>Campaign Name</th><th>Network</th>"
            "<th>Impressions</th><th>Contribution</th></tr>"
            "<tr><td>C1</td><td>CNN</td><td>100</td><td>80</td></tr>"
            "<tr><td>C1</td><td>ESPN</td><td>200</td><td>150</td></tr>"
            "</table></body></html>")
    p = tmp_path / "report.html"
    p.write_text(html)
    from parsers.html_parser import HTMLParser
    data = HTMLParser(str(p), "TestPlat").parse()
    assert "level_data" in data and "campaign_metrics" in data
    imps = [r for r in data["level_data"] if r["metric_name"] == "Impressions"]
    assert sum(r["metric_value"] for r in imps) == 300
    assert all(r["_campaign"] == "C1" for r in data["level_data"])


def test_ratio_metrics_average_instead_of_sum():
    """Regression: 63 daily Frequency rows of ~2.5 summed to a garbage
    159.74 'Total Frequency'. Ratio metrics average (proven via the mapper
    catalog); the REVIEW omits Reach/Frequency entirely since 2026-07-14 —
    non-dedup values were misleading next to vendor dashboards."""
    rows = [{"metric_name": "Frequency", "metric_value": 2.5,
             "metric_level": f"date:D{i}", "_campaign": "C1"} for i in range(63)]
    data = [{"campaign_metrics": {}, "level_data": rows}]
    kpi_totals, details, _ = compute_kpis(data)
    assert not any("Frequency" in k for k in kpi_totals)
    assert "Frequency" not in details.get("C1", {})
    metrics, _ = get_available_metrics(data, "T", "", "")
    assert metrics["Avg Frequency"] == 2.5   # averaging rule unchanged
    assert "Total Frequency" not in metrics


def test_no_float_repr_leak_in_default_format():
    from engine.pptx_formats import format_with_details, _format_value
    noisy = 159.74153592699997
    d = {"format": "text", "decimals": 0, "commas": True, "prefix": "", "suffix": ""}
    assert format_with_details(noisy, d) == "159.74"
    assert _format_value(noisy, "text") == "159.74"


def test_rate_suspicion_flag():
    """Additive metric whose values are all tiny gets flagged as a
    probable mis-aliased rate (the 'Completions: 1.00' mystery)."""
    rows = [{"metric_name": "Completions", "metric_value": v,
             "metric_level": f"date:D{i}", "_campaign": "C1"}
            for i, v in enumerate([0.0, 0.01, 0.94, 0.05, 0.0, 0.0])]
    data = [{"campaign_metrics": {}, "level_data": rows}]
    _, _, flags = compute_kpis(data)
    assert any("rate/percentage" in f for f in flags)


def test_date_format_with_style():
    from engine.pptx_formats import format_date_with_style, format_with_details
    assert format_date_with_style("2026-05-01", "long_ordinal") == "May 1st, 2026"
    assert format_date_with_style("2026-05-01 - 2026-05-31", "us_slash") == \
        "05/01/2026 – 05/31/2026"
    assert format_date_with_style("2026-05-01", "custom", "%b '%y") == "May '26"
    d = {"format": "date", "date_style": "month_year"}
    assert format_with_details("2026-05-01", d) == "May 2026"


def test_structured_breakdown_preview_averages_ratio_metrics():
    rows = [
        {"metric_name": "CTR", "metric_value": 0.10,
         "metric_level": "device:Roku", "_campaign": "C1"},
        {"metric_name": "CTR", "metric_value": 0.30,
         "metric_level": "device:Roku", "_campaign": "C1"},
    ]
    _, structured = get_available_metrics(
        [{"campaign_metrics": {}, "level_data": rows}], "T", "", "")
    assert structured["breakdowns"]["device"][0]["CTR"] == 0.20
