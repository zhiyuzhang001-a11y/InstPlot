#!/bin/sh
set -u

SCRIPT_DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
INSTALLER="$SCRIPT_DIR/scripts/install.py"

run_installer() {
    if command -v python3.12 >/dev/null 2>&1; then
        python3.12 "$INSTALLER" --apply "$@"
    elif command -v python3 >/dev/null 2>&1; then
        python3 "$INSTALLER" --apply "$@"
    elif command -v uv >/dev/null 2>&1; then
        uv run --no-project --python 3.12 "$INSTALLER" --apply "$@"
    else
        printf '%s\n' "未找到 Python。请从 https://www.python.org/ 安装 Python 3.12，或安装 uv 后重试。"
        return 1
    fi
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
