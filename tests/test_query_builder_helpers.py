"""Tests for the query builder's non-UI helpers.

``build_pivot`` feeds tables and charts that ship in client reports, so the
cross-breakdown aggregation rules are locked here: values must never be
pooled across breakdown types (each type repeats the same delivery, and
most types have an "Other" bucket that ballooned into a bogus top row).
"""

import pandas as pd
import pytest


@pytest.fixture
def raw_df():
    """Mirror of the query builder's raw-rows schema: one campaign total
    plus zone/dow/network breakdowns that each re-slice the same delivery.
    'Other' exists in all three types — the shipped-report bug shape."""
    return pd.DataFrame([
        {"campaign": "Campaign A", "metric": "Airings", "value": 1000,
         "source": "campaign", "level_value": ""},
        {"campaign": "Campaign A", "metric": "Airings", "value": 800,
         "source": "zone", "level_value": "Honolulu"},
        {"campaign": "Campaign A", "metric": "Airings", "value": 200,
         "source": "zone", "level_value": "Other"},
        {"campaign": "Campaign A", "metric": "Airings", "value": 600,
         "source": "dow", "level_value": "Monday"},
        {"campaign": "Campaign A", "metric": "Airings", "value": 400,
         "source": "dow", "level_value": "Other"},
        {"campaign": "Campaign A", "metric": "Airings", "value": 1000,
         "source": "network", "level_value": "Other"},
    ])


def _pivot(raw_df, sel_sources, topn="all"):
    from mapper.query_builder import build_pivot
    pivot, _note = build_pivot(raw_df, "Airings", "sum", topn,
                               [], sel_sources, [])
    return pivot


def test_no_breakdown_selected_shows_campaign_totals_not_pooled_levels(raw_df):
    """The photo bug: with no breakdown type selected, every type's rows
    were pooled and 'Other' summed to 200+400+1000. The documented default
    is a clean campaign-totals table."""
    pivot = _pivot(raw_df, sel_sources=[])

    assert list(pivot.index) == ["Campaign A"]
    assert pivot.iloc[0, -1] == 1000
    assert "Other" not in pivot.index


def test_single_breakdown_type_totals_stay_within_that_type(raw_df):
    pivot = _pivot(raw_df, sel_sources=["zone"])

    assert list(pivot.index) == ["Honolulu", "Other"]  # sorted desc by total
    assert pivot.loc["Other", "Total"] == 200          # zone's Other only
    assert pivot["Total"].sum() == 1000


def test_same_value_in_multiple_types_is_disambiguated_not_summed(raw_df):
    pivot = _pivot(raw_df, sel_sources=["zone", "dow"])

    assert "Other" not in pivot.index          # never silently merged
    assert pivot.loc["Other (zone)", "Total"] == 200
    assert pivot.loc["Other (dow)", "Total"] == 400
    assert pivot.loc["Honolulu", "Total"] == 800   # unambiguous stays plain
    assert pivot.loc["Monday", "Total"] == 600


def test_top_n_keeps_largest_rows(raw_df):
    pivot = _pivot(raw_df, sel_sources=["zone", "dow"], topn="2")

    assert len(pivot) == 2
    assert list(pivot.index) == ["Honolulu", "Monday"]


def test_no_matching_rows_returns_none(raw_df):
    from mapper.query_builder import build_pivot
    pivot, note = build_pivot(raw_df, "Clicks", "sum", "all", [], [], [])

    assert pivot is None
    assert note == "No data matches filters"


def test_auto_query_name_is_readable():
    from mapper.query_builder import _auto_query_name

    name = _auto_query_name({"metric": "Impressions", "agg": "sum"}, {})
    assert name == "Query: Impressions (sum)"


def test_auto_query_name_dedupes_against_existing():
    from mapper.query_builder import _auto_query_name

    existing = {"Query: Impressions (sum)": {},
                "Query: Impressions (sum) 2": {}}
    name = _auto_query_name({"metric": "Impressions", "agg": "sum"}, existing)
    assert name == "Query: Impressions (sum) 3"


def test_auto_query_name_survives_missing_keys():
    from mapper.query_builder import _auto_query_name

    assert _auto_query_name({}, {}) == "Query: Value (sum)"
