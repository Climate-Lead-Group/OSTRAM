"""Fail-closed notice for the archived WS-3 in-place template writer."""

from __future__ import annotations

import sys


ARCHIVED_COPY = "docs/archive/ws3-ws4/scripts/set_final_v18_interconnector_values.py"


def main() -> int:
    print(
        "ERROR: This one-shot template writer is disabled. "
        f"Its historical source is preserved at {ARCHIVED_COPY}.",
        file=sys.stderr,
    )
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
