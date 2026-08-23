from pathlib import Path

from scripts.verify_install import run_smoke


ROOT = Path(__file__).resolve().parents[1]


def test_installed_runtime_smoke_covers_io_plot_and_resources():
    result = run_smoke()
    assert result["status"] == "healthy"
    assert result["svg_count"] == 67
    assert result["txt_rows"] == 2
    assert result["exports"] == ["csv", "png", "xlsx"]


def test_platform_install_entrypoints_are_local_and_non_privileged():
    for name in ["install_windows.bat", "install_macos.command", "install_linux.sh"]:
        path = ROOT / name
        content = path.read_text(encoding="utf-8")
        assert "scripts/install.py" in content or "scripts\\install.py" in content
        assert "--apply" in content
        assert "run_instplot" in content
        assert "sudo" not in content.lower()
        assert "curl" not in content.lower()
        assert "powershell" not in content.lower()


def test_unix_install_entrypoints_are_executable():
    for name in ["install_macos.command", "install_linux.sh"]:
        assert (ROOT / name).stat().st_mode & 0o111
