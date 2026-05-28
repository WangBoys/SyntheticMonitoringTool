#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV_PYTHON="${ROOT_DIR}/.venv/bin/python"
CLI_SCRIPT="${ROOT_DIR}/run_cli.py"

run_with_python() {
  local python_bin="$1"
  shift
  # --config 未指定时，run_cli.py 默认使用当前目录下的 config.json
  exec "$python_bin" "$CLI_SCRIPT" "$@"
}

if [[ -x "$VENV_PYTHON" ]]; then
  run_with_python "$VENV_PYTHON" "$@"
fi

if [[ -d "${ROOT_DIR}/vendor" ]]; then
  export PYTHONPATH="${ROOT_DIR}/vendor${PYTHONPATH:+:${PYTHONPATH}}"
  if ! command -v python3 >/dev/null 2>&1; then
    echo "未找到 python3，请先执行 install_server_deps.sh 或安装 Python 3。" >&2
    exit 1
  fi
  run_with_python "$(command -v python3)" "$@"
fi

echo "未找到运行环境：开发机请执行 ./setup_ubuntu.sh；离线包请先解压完整目录。" >&2
exit 1
