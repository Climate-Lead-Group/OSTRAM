"""Profile-aware routing for existing country preparation helpers."""

from __future__ import annotations

import argparse
from contextlib import contextmanager
import importlib
import sys
from typing import Iterator, Sequence


MODULES = {
    "template": "ostram.pipeline.preparation.country_templates",
    "merge": "ostram.pipeline.preparation.merge_country_template",
    "validate": "ostram.pipeline.preparation.country_validation",
}


@contextmanager
def _arguments(action: str, arguments: Sequence[str]) -> Iterator[None]:
    previous = sys.argv
    sys.argv = [f"python -m ostram country {action}", *arguments]
    try:
        yield
    finally:
        sys.argv = previous


def main(argv: Sequence[str] | None = None) -> int:
    arguments = list(sys.argv[1:] if argv is None else argv)
    parser = argparse.ArgumentParser(prog="python -m ostram country", add_help=False)
    parser.add_argument("action", choices=sorted(MODULES))
    if not arguments or arguments == ["--help"] or arguments == ["-h"]:
        print("usage: python -m ostram country {merge,template,validate} ...")
        print("\nprofile-aware country data operations")
        print("\npositional arguments:")
        print("  {merge,template,validate}")
        return 0
    known, remaining = parser.parse_known_args(arguments)
    module = importlib.import_module(MODULES[known.action])
    with _arguments(known.action, remaining):
        result = module.main()
    return result if isinstance(result, int) else 0


if __name__ == "__main__":
    raise SystemExit(main())
