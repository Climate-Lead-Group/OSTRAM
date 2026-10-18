"""Create a derived scenario as a copy of an existing one.

``python -m ostram new-scenario <new> --from <existing>`` copies the folder
``config/scenarios/<existing>/`` to ``config/scenarios/<new>/`` and sets the
``scenario:`` line of the copied ``scenario.yaml`` to ``<new>``. Everything
else in the folder is copied as it is. The command never edits the setup file;
it prints the line to add to its ``decision_scenarios`` list.
"""

from __future__ import annotations

import argparse
import re
import shutil
import sys
from pathlib import Path
from typing import Sequence

from ostram.paths import resolve_paths
from ostram.setups import active_setup_id, load_setup

from .registry import _SCENARIO_NAME, load_registry

_SCENARIO_LINE = re.compile(r"^scenario:[^\r\n]*", re.MULTILINE)


def _relative(path: Path, root: Path) -> str:
    try:
        return path.resolve().relative_to(root.resolve()).as_posix()
    except ValueError:
        return str(path)


def main(argv: Sequence[str] | None = None) -> int:
    """Copy one derived scenario's folder under a new name and print the next steps.

    Returns 1, with the reason on stderr, when the new name is unsafe, already
    used, or the source is not a derived scenario of the active setup.
    """
    parser = argparse.ArgumentParser(
        prog="python -m ostram new-scenario",
        description="Create a derived scenario as a copy of an existing one.",
    )
    parser.add_argument("name", help="name of the new scenario")
    parser.add_argument(
        "--from",
        dest="source",
        required=True,
        help="the derived scenario to copy",
    )
    args = parser.parse_args(argv)
    paths = resolve_paths()
    setup_id = active_setup_id()
    setup_file = _relative(load_setup(setup_id, paths=paths).path, paths.project_root)
    registry = load_registry(paths.scenario_registry)
    folder = paths.scenario_config_root / args.name
    shown = _relative(folder, paths.project_root)

    if not _SCENARIO_NAME.fullmatch(args.name):
        reason = (
            f"{args.name!r} is not a valid scenario name. Use letters, digits, "
            "'_' and '-', starting with a letter or digit."
        )
    elif folder.exists():
        reason = f"{shown} already exists."
    elif args.name in registry.root_names or args.name in registry.decision_scenarios:
        reason = f"{args.name} is already a registered scenario."
    elif args.source not in registry.derived_by_name:
        reason = (
            f"{args.source} is not a derived scenario of the {setup_id} setup. "
            f"Derived scenarios: {', '.join(registry.derived_by_name) or 'none'}."
        )
    else:
        reason = None
    if reason:
        print(f"Cannot create {args.name}: {reason}", file=sys.stderr)
        return 1

    shutil.copytree(paths.scenario_config_root / args.source, folder)
    scenario_file = folder / "scenario.yaml"
    text = scenario_file.read_bytes().decode("utf-8")
    text, count = _SCENARIO_LINE.subn(f"scenario: {args.name}", text, count=1)
    if count != 1:
        shutil.rmtree(folder)
        print(
            f"Cannot create {args.name}: the scenario file of {args.source} "
            "has no 'scenario:' line.",
            file=sys.stderr,
        )
        return 1
    scenario_file.write_bytes(text.encode("utf-8"))

    print(
        f"Created {_relative(scenario_file, paths.project_root)} "
        f"(copied from {args.source}).\n"
        "\n"
        "Next:\n"
        "1. Rewrite the description and edits in that file.\n"
        "2. Add this line at the end of decision_scenarios in "
        f"{setup_file}:\n"
        f"  - {args.name}"
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
