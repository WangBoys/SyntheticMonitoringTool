"""Install and verify Playwright Chromium under pw-browsers for offline bundling."""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
from importlib.metadata import version
from pathlib import Path

META_FILE = ".sync-meta.json"
HEADLESS_BROWSER_NAME = "chromium-headless-shell"


def required_headless_revision() -> str:
    import playwright._impl._driver as driver_mod

    driver_exe = driver_mod.compute_driver_executable()
    driver_path = driver_exe[0] if isinstance(driver_exe, tuple) else driver_exe
    browsers_json = Path(driver_path).parent / "package" / "browsers.json"
    data = json.loads(browsers_json.read_text(encoding="utf-8"))
    for browser in data["browsers"]:
        if browser.get("name") == HEADLESS_BROWSER_NAME:
            return str(browser["revision"])
    raise RuntimeError(f"Missing {HEADLESS_BROWSER_NAME!r} in Playwright browsers.json")


def headless_shell_exe(browsers_path: Path, revision: str) -> Path:
    install_dir = browsers_path / f"chromium_headless_shell-{revision}"
    if sys.platform == "win32":
        return install_dir / "chrome-headless-shell-win64" / "chrome-headless-shell.exe"
    if sys.platform == "linux":
        return install_dir / "chrome-headless-shell-linux64" / "chrome-headless-shell"
    if sys.platform == "darwin":
        return install_dir / "chrome-headless-shell-mac-x64" / "chrome-headless-shell"
    raise RuntimeError(f"Unsupported platform for Playwright headless shell: {sys.platform}")


def load_meta(browsers_path: Path) -> dict[str, str] | None:
    meta_path = browsers_path / META_FILE
    if not meta_path.is_file():
        return None
    try:
        payload = json.loads(meta_path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return None
    if not isinstance(payload, dict):
        return None
    return {str(key): str(value) for key, value in payload.items()}


def playwright_package_version() -> str:
    return version("playwright")


def write_meta(browsers_path: Path, revision: str) -> None:
    meta_path = browsers_path / META_FILE
    meta_path.write_text(
        json.dumps(
            {
                "playwright_version": playwright_package_version(),
                "headless_revision": revision,
            },
            ensure_ascii=True,
            indent=2,
        )
        + "\n",
        encoding="utf-8",
    )


def needs_sync(browsers_path: Path, revision: str, force: bool) -> bool:
    if force:
        return True
    if not headless_shell_exe(browsers_path, revision).is_file():
        return True
    meta = load_meta(browsers_path)
    if meta is None:
        return True
    if meta.get("playwright_version") != playwright_package_version():
        return True
    if meta.get("headless_revision") != revision:
        return True
    return False


def remove_incomplete_headless_install(browsers_path: Path, revision: str) -> None:
    install_dir = browsers_path / f"chromium_headless_shell-{revision}"
    if install_dir.is_dir() and not headless_shell_exe(browsers_path, revision).is_file():
        shutil.rmtree(install_dir, ignore_errors=True)


def prune_stale_browsers(browsers_path: Path, revision: str) -> None:
    for prefix in ("chromium_headless_shell-", "chromium-"):
        for install_dir in browsers_path.glob(f"{prefix}*"):
            if not install_dir.is_dir():
                continue
            if install_dir.name.endswith(f"-{revision}"):
                continue
            shutil.rmtree(install_dir, ignore_errors=True)


def install_chromium(browsers_path: Path) -> None:
    browsers_path.mkdir(parents=True, exist_ok=True)
    env = os.environ.copy()
    env["PLAYWRIGHT_BROWSERS_PATH"] = str(browsers_path.resolve())
    subprocess.run(
        [sys.executable, "-m", "playwright", "install", "chromium"],
        check=True,
        env=env,
    )


def sync_browsers(browsers_path: Path, force: bool = False) -> str:
    revision = required_headless_revision()
    if not needs_sync(browsers_path, revision, force):
        return revision

    remove_incomplete_headless_install(browsers_path, revision)
    install_chromium(browsers_path)
    if not headless_shell_exe(browsers_path, revision).is_file():
        raise RuntimeError(
            "Playwright Chromium install finished but headless shell executable is still missing: "
            f"{headless_shell_exe(browsers_path, revision)}"
        )

    write_meta(browsers_path, revision)
    prune_stale_browsers(browsers_path, revision)
    return revision


def check_browsers(browsers_path: Path) -> str:
    revision = required_headless_revision()
    exe_path = headless_shell_exe(browsers_path, revision)
    if not exe_path.is_file():
        raise RuntimeError(
            f"Missing Playwright Chromium {revision} at {exe_path}. "
            "Run: python pw_browser_sync.py sync"
        )
    return revision


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "command",
        choices=("sync", "check", "revision"),
        help="sync: install if needed; check: verify executable; revision: print required revision",
    )
    parser.add_argument(
        "--path",
        type=Path,
        default=Path(__file__).resolve().parent / "pw-browsers",
        help="Playwright browsers directory (default: project pw-browsers)",
    )
    parser.add_argument("--force", action="store_true", help="Force reinstall for sync")
    args = parser.parse_args(argv)
    browsers_path = args.path.resolve()

    if args.command == "revision":
        print(required_headless_revision())
        return 0
    if args.command == "sync":
        revision = sync_browsers(browsers_path, force=args.force)
        print(revision)
        return 0
    revision = check_browsers(browsers_path)
    print(revision)
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (RuntimeError, subprocess.CalledProcessError) as exc:
        print(exc, file=sys.stderr)
        raise SystemExit(1) from exc
