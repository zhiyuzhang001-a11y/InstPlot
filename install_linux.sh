#!/bin/sh
set -u

SCRIPT_DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
INSTALLER="$SCRIPT_DIR/scripts/install.py"
LOCAL_UV="$SCRIPT_DIR/.installer/uv/uv"
PYTHON_REQUEST=">=3.10,<3.15"

if [ "${INSTPLOT_FORCE_UV_BOOTSTRAP:-0}" != "1" ] && command -v uv >/dev/null 2>&1; then
    UV=$(command -v uv)
else
    "$SCRIPT_DIR/scripts/bootstrap_uv.sh"
    UV="$LOCAL_UV"
fi

"$UV" run --no-project --python "$PYTHON_REQUEST" "$INSTALLER" --apply "$@"

status=$?
if [ "$status" -ne 0 ]; then
    printf '%s\n' "Installation failed. See .install-logs in the project directory."
    exit "$status"
fi

if [ "${INSTPLOT_INSTALL_ONLY:-0}" = "1" ]; then
    exit 0
fi

exec "$SCRIPT_DIR/run_instplot.sh"
