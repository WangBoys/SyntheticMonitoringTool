#!/usr/bin/env bash
# Shared apt package resolution for Playwright/Chromium (Ubuntu 22.04 / 24.04+ t64 variants).

pick_apt_pkg() {
  for candidate in "$@"; do
    if apt-cache show "$candidate" &>/dev/null 2>&1; then
      echo "$candidate"
      return 0
    fi
  done
  echo "警告: 未找到可安装包（已尝试: $*）" >&2
  return 1
}

# Print one package name per line (resolved).
playwright_system_packages() {
  local -a groups=(
    "libnss3"
    "libnspr4"
    "libatk1.0-0 libatk1.0-0t64"
    "libatk-bridge2.0-0 libatk-bridge2.0-0t64"
    "libcups2 libcups2t64"
    "libdrm2"
    "libxkbcommon0"
    "libxcomposite1"
    "libxdamage1"
    "libxfixes3"
    "libxrandr2"
    "libgbm1"
    "libpango-1.0-0"
    "libcairo2"
    "libasound2 libasound2t64"
    "libatspi2.0-0 libatspi2.0-0t64"
  )
  local group pkg
  for group in "${groups[@]}"; do
    # shellcheck disable=SC2086
    if pkg="$(pick_apt_pkg $group)"; then
      echo "$pkg"
    fi
  done
}

install_playwright_system_packages() {
  local run_as_root=("$@")
  if [[ ${#run_as_root[@]} -eq 0 ]]; then
    run_as_root=(sudo)
  fi
  local -a pkgs=()
  while IFS= read -r line; do
    [[ -n "$line" ]] && pkgs+=("$line")
  done < <(playwright_system_packages)
  if [[ ${#pkgs[@]} -eq 0 ]]; then
    echo "未解析到 Playwright 系统库包名。" >&2
    return 1
  fi
  "${run_as_root[@]}" apt-get update
  "${run_as_root[@]}" apt-get install -y "${pkgs[@]}"
}
