"""Fork Stage 0 workflow/CLI and module-boundary checks."""

from __future__ import annotations

import json
import sys

import pytest

from conftest import PROJECT_ROOT

if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)


def test_stage_scaffolds_have_stable_public_boundaries():
    from engine import campaign_dictionary, staging, validate, workflow

    assert callable(campaign_dictionary.apply)
    assert callable(staging.write_staging_workbook)
    assert callable(staging.read_staging_workbook)
    assert callable(validate.validate_ingest)
    assert callable(validate.validate_fill)
    assert callable(workflow.parse_dump)
    assert callable(workflow.generate_staging)
    assert callable(workflow.build_deck)


def test_future_stage_calls_fail_loudly_not_silently():
    from engine import workflow

    with pytest.raises(NotImplementedError, match="Stage 2"):
        workflow.generate_staging("synthetic.xlsx")
    with pytest.raises(NotImplementedError, match="Stage 4"):
        workflow.build_deck()


def test_cli_list_templates_is_json_and_successful(capsys):
    from app import cli

    code = cli.main(["list-templates"])
    payload = json.loads(capsys.readouterr().out)
    assert code == 0
    assert payload["ok"] is True
    assert isinstance(payload["data"], list)
    assert any(item["filename"] == "StandIn_Report.pptx"
               for item in payload["data"])


def test_cli_is_a_thin_workflow_shell():
    import inspect
    from app import cli
    from engine import workflow

    source = inspect.getsource(cli)
    assert "from engine import workflow" in source
    for name in ("parse_dump", "generate_staging", "build_deck",
                 "list_templates", "get_settings", "set_settings",
                 "describe_state"):
        assert hasattr(workflow, name)
