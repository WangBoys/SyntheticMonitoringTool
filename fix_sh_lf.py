from __future__ import annotations

import sys
from pathlib import Path


def normalize_lf(path: Path) -> None:
    text = path.read_text(encoding="utf-8")
    path.write_bytes(text.replace("\r\n", "\n").replace("\r", "\n").encode("utf-8"))


def main(argv: list[str]) -> int:
    if len(argv) < 2:
        print("usage: fix_sh_lf.py <directory>", file=sys.stderr)
        return 1
    directory = Path(argv[1])
    if not directory.is_dir():
        print(f"not a directory: {directory}", file=sys.stderr)
        return 1
    for sh_file in directory.glob("*.sh"):
        normalize_lf(sh_file)
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
