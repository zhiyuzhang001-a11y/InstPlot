@echo off
setlocal
chcp 65001 >nul
set "SCRIPT_DIR=%~dp0"
set "INSTALLER=%SCRIPT_DIR%scripts\install.py"

py -3.12 --version >nul 2>&1
if not errorlevel 1 (
    py -3.12 "%INSTALLER%" --apply %*
    goto after_install
)

where python3 >nul 2>&1
if not errorlevel 1 (
    python3 "%INSTALLER%" --apply %*
    goto after_install
)

where python >nul 2>&1
if not errorlevel 1 (
    python "%INSTALLER%" --apply %*
    goto after_install
)

where uv >nul 2>&1
if not errorlevel 1 (
    uv run --no-project --python 3.12 "%INSTALLER%" --apply %*
    goto after_install
)

echo Python was not found. Install Python 3.12 from https://www.python.org/ or install uv, then retry.
if not "%INSTPLOT_INSTALL_ONLY%"=="1" pause
exit /b 1

:after_install
if errorlevel 1 (
    echo Installation failed. See .install-logs in the project directory.
    if not "%INSTPLOT_INSTALL_ONLY%"=="1" pause
    exit /b 1
)

if "%INSTPLOT_INSTALL_ONLY%"=="1" exit /b 0

call "%SCRIPT_DIR%run_instplot.bat"
