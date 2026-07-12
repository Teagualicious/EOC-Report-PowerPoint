"""Smoke test: proves the test suite and CI pipeline are wired up.

Replace with real tests as the project grows. Never delete the suite itself.
"""


def test_pipeline_is_alive():
    assert True


def test_fixtures_dir_exists():
    from pathlib import Path
    assert (Path(__file__).parent / "fixtures").is_dir()
