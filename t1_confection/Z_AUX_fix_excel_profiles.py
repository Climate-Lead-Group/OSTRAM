"""Fail-closed notice for the archived demand-profile workbook writer."""

from __future__ import annotations

import sys


ARCHIVED_COPY = "docs/archive/legacy-tools/Z_AUX_fix_excel_profiles.py"


def main() -> int:
    print(
        "ERROR: This stale workbook-writing utility is disabled. "
        f"Its historical source is preserved at {ARCHIVED_COPY}.",
        file=sys.stderr,
    )
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
