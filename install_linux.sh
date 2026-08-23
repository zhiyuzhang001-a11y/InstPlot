#!/bin/sh
set -u

SCRIPT_DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
INSTALLER="$SCRIPT_DIR/scripts/install.py"

if command -v python3.12 >/dev/null 2>&1; then
    python3.12 "$INSTALLER" --apply "$@"
elif command -v python3 >/dev/null 2>&1; then
    python3 "$INSTALLER" --apply "$@"
elif command -v uv >/dev/null 2>&1; then
    uv run --no-project --python 3.12 "$INSTALLER" --apply "$@"
else
    printf '%s\n' "Python was not found. Install Python 3.12 from python.org or install uv, then retry."
    exit 1
fi

status=$?
if [ "$status" -ne 0 ]; then
    printf '%s\n' "Installation failed. See .install-logs in the project directory."
    exit "$status"
fi

if [ "${INSTPLOT_INSTALL_ONLY:-0}" = "1" ]; then
    exit 0
fi

exec "$SCRIPT_DIR/run_instplot.sh"
