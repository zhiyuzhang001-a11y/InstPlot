import os
from pathlib import Path
from types import SimpleNamespace

import scripts.verify_release as release
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
    assert workflow.index("--measure-full-pyside") < workflow.index(
        "-m pip install pytest==8.4.2"
    )


def test_full_pyside_measurement_uses_disposable_environment(tmp_path, monkeypatch):
    environment = tmp_path / ".venv"
    python = release._environment_python(environment)
    python.parent.mkdir(parents=True)
    python.touch()
    sizes = iter((100, 250))
    commands = []

    def fake_run(command, **kwargs):
        commands.append(command)
        if "import importlib.metadata" in " ".join(command):
            return SimpleNamespace(returncode=0, stdout="6.9.1\n", stderr="")
        return SimpleNamespace(returncode=0, stdout="", stderr="")

    monkeypatch.setattr(release, "environment_size", lambda ignored: next(sizes))
    monkeypatch.setattr(release.subprocess, "run", fake_run)

    result = release.measure_full_pyside(environment)

    assert result["essentials_bytes"] == 100
    assert result["full_bytes"] == 250
    assert all(command[0] != str(python) for command in commands)
    assert not any("uninstall" in command for command in commands)
