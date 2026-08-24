#!/bin/sh
set -u

SCRIPT_DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
INSTALLER="$SCRIPT_DIR/scripts/install.py"
LOCAL_UV="$SCRIPT_DIR/.installer/uv/uv"
PYTHON_REQUEST=">=3.10,<3.15"

run_installer() {
    if [ "${INSTPLOT_FORCE_UV_BOOTSTRAP:-0}" != "1" ] && command -v uv >/dev/null 2>&1; then
        UV=$(command -v uv)
    else
        "$SCRIPT_DIR/scripts/bootstrap_uv.sh" || return 1
        UV="$LOCAL_UV"
    fi
    "$UV" run --no-project --python "$PYTHON_REQUEST" "$INSTALLER" --apply "$@"
}

if ! run_installer "$@"; then
    printf '%s\n' "安装失败。请查看项目 .install-logs 目录中的日志。"
    if [ "${INSTPLOT_INSTALL_ONLY:-0}" != "1" ]; then
        printf '%s' "按回车键关闭…"
        read -r _answer
    fi
    exit 1
fi

if [ "${INSTPLOT_INSTALL_ONLY:-0}" = "1" ]; then
    exit 0
fi

exec "$SCRIPT_DIR/run_instplot.command"
