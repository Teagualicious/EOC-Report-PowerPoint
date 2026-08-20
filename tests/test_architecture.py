"""Architecture laws established before any Deck Engine feature work."""

from __future__ import annotations

import ast
import os
import subprocess
import sys
from pathlib import Path

from conftest import APP_DIR, PROJECT_ROOT

APP = Path(APP_DIR)


def _python_files():
    yield from APP.rglob("*.py")


def test_arch_1_business_paths_never_depend_on_cwd():
    """T-ARCH-1: program paths are anchored by config.paths, never the shell."""
    violations = []
    for path in _python_files():
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        for node in ast.walk(tree):
            if not isinstance(node, ast.Call):
                continue
            if (isinstance(node.func, ast.Attribute)
                    and isinstance(node.func.value, ast.Name)
                    and node.func.value.id == "os"
                    and node.func.attr == "getcwd"):
                violations.append(f"{path.relative_to(APP)}:{node.lineno}: os.getcwd")
            if isinstance(node.func, ast.Name) and node.func.id == "Path":
                if not node.args:
                    violations.append(f"{path.relative_to(APP)}:{node.lineno}: Path()")
                elif (isinstance(node.args[0], ast.Constant)
                      and node.args[0].value in ("", ".")):
                    violations.append(f"{path.relative_to(APP)}:{node.lineno}: Path('.')")
    assert violations == []


def test_arch_2_config_naming_is_the_only_filename_sanitizer():
    """T-ARCH-2: user-controlled path-component logic has one authority."""
    protected = {"safe_component", "storage_key", "safe_child", "unique_component",
                 "sanitize_filename", "safe_filename"}
    violations = []
    for path in _python_files():
        if path == APP / "config" / "naming.py":
            continue
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        for node in ast.walk(tree):
            if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)) \
                    and node.name in protected:
                violations.append(f"{path.relative_to(APP)}:{node.lineno}: {node.name}")
    assert violations == []


def test_arch_3_headless_engine_imports_do_not_reach_tk():
    """T-ARCH-3: workflow/staging/validation/fill import without tkinter."""
    code = """
import json, sys
from engine import workflow, staging, validate, pptx_fill
print(json.dumps(sorted(name for name in sys.modules if name.startswith('tkinter'))))
"""
    env = dict(os.environ)
    env["PYTHONPATH"] = APP_DIR + os.pathsep + env.get("PYTHONPATH", "")
    result = subprocess.run(
        [sys.executable, "-c", code], cwd=PROJECT_ROOT, env=env,
        check=True, capture_output=True, text=True,
    )
    assert result.stdout.strip() == "[]"


def test_stripped_runtime_modules_are_absent():
    """T-ARCH-4: removed Stage 0 runtime modules must stay absent."""
    stripped = [
        APP / "mcp_server.py",
        APP / "engine" / "excel_writer.py",
        APP / "engine" / "excel_search_dashboard.py",
        APP / "engine" / "excel_vba.py",
        APP / "engine" / "pptx_live.py",
        APP / "ui" / "main_window.py",
        APP / "ui" / "review_view.py",
        APP / "ui" / "client_wizard.py",
        APP / "ui" / "settings_window.py",
    ]
    assert [str(path.relative_to(APP)) for path in stripped if path.exists()] == []
