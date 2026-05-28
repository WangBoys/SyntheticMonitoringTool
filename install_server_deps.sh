#!/usr/bin/env bash
# 服务器首次部署时执行一次（需 sudo），安装 Playwright/Chromium 所需的系统库。
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
# shellcheck source=ubuntu_apt_deps.sh
source "${ROOT_DIR}/ubuntu_apt_deps.sh"

if [[ "$(id -u)" -ne 0 ]] && ! command -v sudo >/dev/null 2>&1; then
  echo "请使用 root 或 sudo 运行此脚本。" >&2
  exit 1
fi

run_as_root() {
  if [[ "$(id -u)" -eq 0 ]]; then
    "$@"
  else
    sudo "$@"
  fi
}

if ! command -v apt-get >/dev/null 2>&1; then
  echo "未找到 apt-get，请手动安装 Playwright 系统依赖：" >&2
  echo "  python3 -m playwright install-deps chromium" >&2
  exit 1
fi

echo "安装 Playwright/Chromium 系统依赖（兼容 Ubuntu 24.04 t64 包名）..."
run_as_root apt-get update
run_as_root apt-get install -y python3
mapfile -t PW_PKGS < <(playwright_system_packages)
run_as_root apt-get install -y "${PW_PKGS[@]}"

echo "系统依赖安装完成。"
echo "若已创建 .venv，也可在项目中执行: python -m playwright install-deps chromium"
