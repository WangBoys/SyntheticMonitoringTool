from __future__ import annotations

import os
import sys


def runtime_base_dir() -> str:
    if getattr(sys, "frozen", False):
        return getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.dirname(os.path.abspath(__file__))


def setup_runtime_environment() -> str:
    base_dir = runtime_base_dir()
    exe_dir = os.path.dirname(os.path.abspath(sys.executable)) if getattr(sys, "frozen", False) else base_dir
    runtime_root = os.path.join(base_dir, "runtime-data")
    runtime_temp = os.path.join(runtime_root, "temp")
    runtime_cache = os.path.join(runtime_root, "cache")
    os.makedirs(runtime_temp, exist_ok=True)
    os.makedirs(runtime_cache, exist_ok=True)

    os.environ["TMP"] = runtime_temp
    os.environ["TEMP"] = runtime_temp
    os.environ["TMPDIR"] = runtime_temp
    os.environ["XDG_CACHE_HOME"] = runtime_cache
    os.environ["PYTHON_EGG_CACHE"] = os.path.join(runtime_cache, "python-eggs")

    bundled_browser_paths = [
        os.path.join(base_dir, "pw-browsers"),
        os.path.join(exe_dir, "pw-browsers"),
    ]
    for bundled_browser_path in bundled_browser_paths:
        if os.path.isdir(bundled_browser_path):
            os.environ["PLAYWRIGHT_BROWSERS_PATH"] = bundled_browser_path
            break
    else:
        os.environ.setdefault("PLAYWRIGHT_BROWSERS_PATH", "0")

    return base_dir
