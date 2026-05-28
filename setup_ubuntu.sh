#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$ROOT_DIR"
# shellcheck source=ubuntu_apt_deps.sh
source "${ROOT_DIR}/ubuntu_apt_deps.sh"

echo "[1/5] Installing base system packages (requires sudo)..."
if command -v apt-get >/dev/null 2>&1; then
  sudo apt-get update
  sudo apt-get install -y python3 python3-venv python3-pip
else
  echo "apt-get not found; please install Python 3 manually." >&2
  exit 1
fi

echo "[2/5] Creating virtual environment..."
python3 -m venv .venv
# shellcheck disable=SC1091
source .venv/bin/activate

echo "[3/5] Installing Python dependencies..."
python -m pip install --upgrade pip
python -m pip install -r requirements-linux.txt

echo "[4/5] Installing Playwright Chromium system deps and browser..."
python -m playwright install-deps chromium
python pw_browser_sync.py sync

echo "[5/5] Preparing config template..."
if [[ ! -f config.json ]]; then
  cp config.example.json config.json
  echo "Created config.json from config.example.json — please edit it before running."
else
  echo "config.json already exists, skipped copy."
fi

chmod +x run_cli.sh

echo
echo "Setup complete. Next steps:"
echo "  vim config.json"
echo "  ./run_cli.sh --dry-run"
echo "  ./run_cli.sh"
echo "  # 或指定配置: ./run_cli.sh --config /path/to/other.json"
echo
echo "Tip: run_cli.sh 会自动使用 .venv，无需手动 source activate。"
