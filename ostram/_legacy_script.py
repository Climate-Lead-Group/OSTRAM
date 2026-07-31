"""Execute one staged legacy Python file behind a ``python -m`` boundary.

This adapter exists only while Stage 11 keeps production files in their Pull
Request A locations. Stage 12 replaces these file targets with package modules,
and Stage 13 removes the adapter.
"""

from __future__ import annotations

import argparse
from pathlib import Path
import runpy
import sys


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--script", required=True, type=Path)
    arguments, remainder = parser.parse_known_args(argv)
    script = arguments.script.expanduser().resolve()
    if not script.is_file():
        raise FileNotFoundError(f"legacy Python stage not found: {script}")
    if remainder[:1] == ["--"]:
        remainder = remainder[1:]
    previous = sys.argv
    sys.argv = [script.name, *remainder]
    try:
        runpy.run_path(str(script), run_name="__main__")
    finally:
        sys.argv = previous
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
