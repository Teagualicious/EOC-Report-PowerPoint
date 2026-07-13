"""Tests for the query builder's non-UI helpers."""


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
