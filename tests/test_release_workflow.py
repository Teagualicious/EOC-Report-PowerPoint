"""Static contract checks for the push-driven release workflow."""

from pathlib import Path


REPOSITORY_ROOT = Path(__file__).resolve().parents[1]


def test_release_workflow_uses_versioned_tag_and_archive_names():
    workflow = (REPOSITORY_ROOT / ".github" / "workflows" / "release.yml").read_text(
        encoding="utf-8"
    )

    assert "tags:" in workflow
    assert '"v*.*.*"' in workflow
    assert 'TAG="v$RELEASE_VERSION"' in workflow
    assert 'echo "tag=$TAG"' in workflow
    assert "Deck-Engine-v$VER.zip" in workflow
    assert "Deck-Engine-v$VER-portable-win64.zip" in workflow
    assert '--title "Deck-Engine-v$VER"' in workflow


def test_release_workflow_refuses_development_versions():
    workflow = (REPOSITORY_ROOT / ".github" / "workflows" / "release.yml").read_text(
        encoding="utf-8"
    )

    assert '[[ "$RELEASE_VERSION" == *dev* ]]' in workflow
    assert "publish=false" in workflow
    assert "Development VERSION detected; no release will be published." in workflow
