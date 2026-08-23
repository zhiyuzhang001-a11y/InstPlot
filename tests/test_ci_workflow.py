from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def test_native_install_workflow_covers_all_platforms_and_repair():
    workflow = (ROOT / ".github" / "workflows" / "native-install.yml").read_text(
        encoding="utf-8"
    )
    for runner in ["ubuntu-latest", "macos-latest", "windows-latest"]:
        assert runner in workflow
    for entrypoint in ["install_linux.sh", "install_macos.command", "install_windows.bat"]:
        assert entrypoint in workflow
    assert "actions/checkout@v6" in workflow
    assert "actions/setup-python@v6" in workflow
    assert "contents: read" in workflow
    assert "pull_request_target" not in workflow
    assert "--repair" in workflow
    assert "verify_install.py" in workflow
    assert "pytest" in workflow


def test_entrypoints_support_nonlaunching_ci_mode():
    for name in ["install_windows.bat", "install_macos.command", "install_linux.sh"]:
        content = (ROOT / name).read_text(encoding="utf-8")
        assert "INSTPLOT_INSTALL_ONLY" in content
