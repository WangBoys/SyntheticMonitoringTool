#!/usr/bin/env bash
# 在 Ubuntu/Debian 上构建可拷贝到服务器的离线包（含 Python 依赖与 Playwright 浏览器）。
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$ROOT_DIR"

PACKAGE_NAME="SyntheticMonitoringTool-linux"
STAGING_DIR="${ROOT_DIR}/dist/${PACKAGE_NAME}"
ARCHIVE_PATH="${ROOT_DIR}/dist/${PACKAGE_NAME}.tar.gz"

CLI_FILES=(
  probe_core.py
  config_loader.py
  runtime_env.py
  run_cli.py
  run_cli.sh
  install_server_deps.sh
  config.example.json
  pw_browser_sync.py
)

require_linux() {
  if [[ "$(uname -s)" != "Linux" ]]; then
    echo "此脚本需在 Linux（推荐 Ubuntu）上运行。" >&2
    exit 1
  fi
}

require_python3() {
  if ! command -v python3 >/dev/null 2>&1; then
    echo "未找到 python3，请先安装 Python 3。" >&2
    exit 1
  fi
}

write_package_readme() {
  cat > "${STAGING_DIR}/DEPLOY.md" <<'EOF'
# SyntheticMonitoringTool Linux 离线包部署说明

## 1. 服务器首次准备（只需一次）

解压后进入目录，安装系统库（需 sudo）：

```bash
sudo ./install_server_deps.sh
```

## 2. 配置

```bash
cp config.example.json config.json
vim config.json
```

## 3. 运行

```bash
./run_cli.sh --dry-run
./run_cli.sh
```

未指定 `--config` 时默认读取当前目录下的 `config.json`。

## 说明

- 本包已内置 `vendor/` Python 依赖与 `pw-browsers/` Chromium，无需 pip/venv。
- 服务器需已安装 `python3`（3.10+ 推荐）及 `install_server_deps.sh` 中的系统库。
- 运行时数据写入 `runtime-data/`。
EOF
}

main() {
  require_linux
  require_python3

  echo "[1/6] 准备打包目录..."
  rm -rf "$STAGING_DIR"
  mkdir -p "$STAGING_DIR/vendor" "${STAGING_DIR}/pw-browsers"

  echo "[2/6] 复制程序文件..."
  for file in "${CLI_FILES[@]}"; do
    if [[ ! -f "${ROOT_DIR}/${file}" ]]; then
      echo "缺少文件: ${file}" >&2
      exit 1
    fi
    cp "${ROOT_DIR}/${file}" "${STAGING_DIR}/${file}"
  done
  chmod +x "${STAGING_DIR}/run_cli.sh" "${STAGING_DIR}/install_server_deps.sh"

  echo "[3/6] 安装 Python 依赖到 vendor/..."
  python3 -m pip install --upgrade pip
  python3 -m pip install -r "${ROOT_DIR}/requirements-linux.txt" -t "${STAGING_DIR}/vendor" --upgrade

  echo "[4/6] 同步 Playwright Chromium 到 pw-browsers/..."
  PLAYWRIGHT_BROWSERS_PATH="${STAGING_DIR}/pw-browsers" \
    python3 "${ROOT_DIR}/pw_browser_sync.py" sync --path "${STAGING_DIR}/pw-browsers"

  echo "[5/6] 写入部署说明..."
  write_package_readme
  date -u +"%Y-%m-%dT%H:%M:%SZ" > "${STAGING_DIR}/BUILD_TIME.txt"
  python3 --version > "${STAGING_DIR}/BUILD_PYTHON.txt"

  echo "[6/6] 生成压缩包..."
  mkdir -p "${ROOT_DIR}/dist"
  tar -czf "$ARCHIVE_PATH" -C "${ROOT_DIR}/dist" "$PACKAGE_NAME"

  echo
  echo "打包完成:"
  echo "  目录: ${STAGING_DIR}"
  echo "  压缩包: ${ARCHIVE_PATH}"
  echo
  echo "部署到服务器:"
  echo "  scp ${ARCHIVE_PATH} user@server:/opt/"
  echo "  ssh user@server 'cd /opt && tar -xzf ${PACKAGE_NAME}.tar.gz'"
  echo "  ssh user@server 'cd /opt/${PACKAGE_NAME} && sudo ./install_server_deps.sh'"
  echo "  ssh user@server 'cd /opt/${PACKAGE_NAME} && cp config.example.json config.json && ./run_cli.sh --dry-run'"
}

main "$@"
