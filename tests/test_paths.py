"""Regression tests for the cleaned app/workspace folder layout."""

import json
import os

from config import paths


def _redirect_layout(monkeypatch, tmp_path):
    project = tmp_path / "project"
    workspace = project / "workspace"
    values = {
        "PROJECT_ROOT": str(project),
        "APP_ROOT": str(project),
        "WORKSPACE_DIR": str(workspace),
        "SETTINGS_PATH": str(workspace / "settings.json"),
        "MAPPINGS_DIR": str(workspace / "mappings"),
        "TEMPLATES_DIR": str(workspace / "templates"),
        "IMAGES_DIR": str(workspace / "templates" / "images"),
        "INPUT_DIR": str(project / "input"),
        "OUTPUT_DIR": str(project / "output"),
        "LOGS_DIR": str(workspace / "logs"),
    }
    for name, value in values.items():
        monkeypatch.setattr(paths, name, value)
    return project, workspace


def test_source_layout_points_to_app_resources_and_root_user_folders():
    assert os.path.basename(paths.APP_DIR) == "app"
    assert paths.PROJECT_ROOT == os.path.dirname(paths.APP_DIR)
    assert paths.RESOURCE_DIR == os.path.join(paths.APP_DIR, "resources")
    assert paths.WORKSPACE_DIR == os.path.join(paths.PROJECT_ROOT, "workspace")
    assert paths.INPUT_DIR == os.path.join(paths.PROJECT_ROOT, "input")
    assert paths.OUTPUT_DIR == os.path.join(paths.PROJECT_ROOT, "output")
    assert os.path.isfile(paths.DICT_PATH)
    assert os.path.isfile(paths.LOGO_PATH)


def test_ensure_dirs_creates_workspace_and_root_user_folders(monkeypatch, tmp_path):
    project, workspace = _redirect_layout(monkeypatch, tmp_path)
    paths.ensure_dirs()
    for relative in ("mappings", "templates", "templates/images", "logs"):
        assert (workspace / relative).is_dir()
    assert (project / "input").is_dir()
    assert (project / "output").is_dir()


def test_legacy_migration_is_copy_only_and_does_not_overwrite(monkeypatch, tmp_path):
    project, workspace = _redirect_layout(monkeypatch, tmp_path)

    # Original pre-cleanup root-level application state.
    (project / "templates").mkdir(parents=True)
    (project / "mappings").mkdir()
    (project / "templates" / "legacy.pptx").write_bytes(b"legacy-template")
    (project / "mappings" / "platform_legacy.json").write_text(
        '{"legacy": true}', encoding="utf-8")
    (project / "settings.json").write_text(
        json.dumps({"theme": "dark"}), encoding="utf-8")
    (project / "input_files").mkdir()
    (project / "input_files" / "older.csv").write_text(
        "campaign,impressions\nA,1\n", encoding="utf-8")

    # Previous cleaned build stored input/output under workspace/.
    (workspace / "input").mkdir(parents=True)
    (workspace / "output" / "Client A").mkdir(parents=True)
    (workspace / "input" / "recent.csv").write_text(
        "campaign,impressions\nB,2\n", encoding="utf-8")
    (workspace / "output" / "Client A" / "report.xlsx").write_bytes(
        b"legacy-report")

    # Existing destinations always win.
    (workspace / "templates").mkdir(parents=True)
    (workspace / "templates" / "legacy.pptx").write_bytes(b"newer-template")
    (project / "output" / "Client A").mkdir(parents=True)
    (project / "output" / "Client A" / "report.xlsx").write_bytes(
        b"newer-report")

    paths.ensure_dirs()

    assert (workspace / "templates" / "legacy.pptx").read_bytes() == b"newer-template"
    assert (workspace / "mappings" / "platform_legacy.json").is_file()
    assert json.loads((workspace / "settings.json").read_text(encoding="utf-8"))["theme"] == "dark"
    assert (project / "input" / "older.csv").is_file()
    assert (project / "input" / "recent.csv").is_file()
    assert (project / "output" / "Client A" / "report.xlsx").read_bytes() == b"newer-report"

    # Migration is non-destructive.
    assert (project / "templates" / "legacy.pptx").is_file()
    assert (project / "settings.json").is_file()
    assert (workspace / "input" / "recent.csv").is_file()
    assert (workspace / "output" / "Client A" / "report.xlsx").is_file()


def test_previous_default_output_setting_moves_to_root(monkeypatch, tmp_path):
    from config import settings

    project, workspace = _redirect_layout(monkeypatch, tmp_path)
    monkeypatch.setattr(settings, "SETTINGS_PATH", str(workspace / "settings.json"))
    monkeypatch.setattr(settings, "WORKSPACE_DIR", str(workspace))
    monkeypatch.setattr(settings, "OUTPUT_DIR", str(project / "output"))
    workspace.mkdir(parents=True)
    (workspace / "settings.json").write_text(
        json.dumps({"theme": "dark", "output_folder": str(workspace / "output")}),
        encoding="utf-8",
    )

    loaded = settings.load_settings()
    assert loaded["output_folder"] == str(project / "output")
    assert loaded["theme"] == "dark"


def test_custom_output_setting_is_preserved(monkeypatch, tmp_path):
    from config import settings

    project, workspace = _redirect_layout(monkeypatch, tmp_path)
    custom = project / "custom-reports"
    monkeypatch.setattr(settings, "SETTINGS_PATH", str(workspace / "settings.json"))
    monkeypatch.setattr(settings, "WORKSPACE_DIR", str(workspace))
    monkeypatch.setattr(settings, "OUTPUT_DIR", str(project / "output"))
    workspace.mkdir(parents=True)
    (workspace / "settings.json").write_text(
        json.dumps({"output_folder": str(custom)}), encoding="utf-8")

    assert settings.load_settings()["output_folder"] == str(custom)


def test_renamed_project_stale_default_output_falls_back(monkeypatch, tmp_path):
    """Regression: opening Settings once pinned the ABSOLUTE default output
    path into settings.json; after the project folder was renamed, exports
    recreated the old folder name and silently wrote there. A default-shaped
    path (…/output) that no longer exists follows the current project."""
    from config import settings

    project, workspace = _redirect_layout(monkeypatch, tmp_path)
    monkeypatch.setattr(settings, "SETTINGS_PATH", str(workspace / "settings.json"))
    monkeypatch.setattr(settings, "WORKSPACE_DIR", str(workspace))
    monkeypatch.setattr(settings, "OUTPUT_DIR", str(project / "output"))
    workspace.mkdir(parents=True)
    old_default = tmp_path / "Jughead-Data-Engine-1.27.0" / "output"  # never created
    (workspace / "settings.json").write_text(
        json.dumps({"output_folder": str(old_default)}), encoding="utf-8")

    assert settings.load_settings()["output_folder"] == str(project / "output")


def test_existing_folder_named_output_is_a_deliberate_choice(monkeypatch, tmp_path):
    """A custom folder that happens to be named 'output' but still exists
    on disk must NOT be treated as a stale default."""
    from config import settings

    project, workspace = _redirect_layout(monkeypatch, tmp_path)
    monkeypatch.setattr(settings, "SETTINGS_PATH", str(workspace / "settings.json"))
    monkeypatch.setattr(settings, "WORKSPACE_DIR", str(workspace))
    monkeypatch.setattr(settings, "OUTPUT_DIR", str(project / "output"))
    workspace.mkdir(parents=True)
    elsewhere = tmp_path / "network-share" / "output"
    elsewhere.mkdir(parents=True)
    (workspace / "settings.json").write_text(
        json.dumps({"output_folder": str(elsewhere)}), encoding="utf-8")

    assert settings.load_settings()["output_folder"] == str(elsewhere)
