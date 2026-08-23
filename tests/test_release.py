import os
from pathlib import Path

from scripts.verify_release import environment_size, validate_release_docs


ROOT = Path(__file__).resolve().parents[1]


def test_release_documentation_has_no_placeholders_and_matches_metadata():
    assert validate_release_docs(ROOT) == []


def test_environment_size_counts_regular_files_without_following_symlinks(tmp_path):
    environment = tmp_path / "environment"
    environment.mkdir()
    (environment / "first.bin").write_bytes(b"1234")
    nested = environment / "nested"
    nested.mkdir()
    (nested / "second.bin").write_bytes(b"123456")
    outside = tmp_path / "outside.bin"
    outside.write_bytes(b"12345678")
    if os.name != "nt":
        (environment / "outside-link").symlink_to(outside)

    assert environment_size(environment) == 10


def test_release_workflow_checks_lock_and_cross_platform_footprint():
    workflow = (ROOT / ".github" / "workflows" / "native-install.yml").read_text(
        encoding="utf-8"
    )
    assert "verify_release.py --check-docs --check-lock" in workflow
    assert "uv==0.12.5" in workflow
    assert "--measure-full-pyside" in workflow
