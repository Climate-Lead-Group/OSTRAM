"""Seed every root scenario from ready-made A-O workbooks.

A setup with the policy ``prepared_root_workbooks: true`` ships its root
A-O workbooks and its A-Xtra workbooks as inputs instead of building them
from OSeMOSYS Global CSVs. For such a setup this step replaces A1 and A2:
it copies the four A-O workbooks into ``A1_Outputs_<root>`` and into the
``_post_a2_snapshot_<root>`` that A3 restores, for every root in the
registry, and copies the A-Xtra workbooks into the extra-inputs folder that
B1 reads.

Usage:
    python -m ostram.pipeline.preparation.prepared_roots
"""

from __future__ import annotations

from pathlib import Path
import shutil
from typing import Iterable, Sequence

from ostram.paths import resolve_paths


POLICY = "prepared_root_workbooks"
ROOT_WORKBOOKS = (
    "A-O_AR_Model_Base_Year.xlsx",
    "A-O_AR_Projections.xlsx",
    "A-O_Demand.xlsx",
    "A-O_Parametrization.xlsx",
)


def seed_root_snapshots(
    root_workbooks: Path,
    extra_inputs: Path,
    a1_outputs: Path,
    extra_target: Path,
    roots: Iterable[str],
    workbooks: Sequence[str] = ROOT_WORKBOOKS,
) -> list[Path]:
    """Copy the root workbooks into each root's output and snapshot folders.

    Args:
        root_workbooks: folder that holds every name in ``workbooks``.
        extra_inputs: folder of ``A-Xtra_*.xlsx`` workbooks.
        a1_outputs: the preparation ``A1_Outputs`` folder.
        extra_target: the preparation extra-inputs folder B1 reads.
        roots: root scenario names.
        workbooks: workbook file names to copy.

    Returns:
        The folders written, in order.

    Raises:
        FileNotFoundError: if a workbook or the extra-inputs folder is missing.

    Side effects:
        Replaces ``A1_Outputs_<root>`` and ``_post_a2_snapshot_<root>`` for
        every root, and copies the A-Xtra workbooks into ``extra_target``.
    """
    missing = [name for name in workbooks if not (root_workbooks / name).is_file()]
    if missing:
        raise FileNotFoundError(f"root workbooks missing from {root_workbooks}: {missing}")
    if not extra_inputs.is_dir():
        raise FileNotFoundError(f"extra-inputs folder not found: {extra_inputs}")
    written: list[Path] = []
    for root in roots:
        for folder in (a1_outputs / f"A1_Outputs_{root}", a1_outputs / f"_post_a2_snapshot_{root}"):
            if folder.exists():
                shutil.rmtree(folder)
            folder.mkdir(parents=True)
            for name in workbooks:
                shutil.copy2(root_workbooks / name, folder / name)
            written.append(folder)
    extra_target.mkdir(parents=True, exist_ok=True)
    for workbook in sorted(extra_inputs.glob("A-Xtra_*.xlsx")):
        shutil.copy2(workbook, extra_target / workbook.name)
    written.append(extra_target)
    return written


def main() -> int:
    """Seed every registered root of the active setup from its root workbooks."""
    from ostram.pipeline.scenarios.registry import load_registry

    paths = resolve_paths()
    roots = load_registry().root_names
    written = seed_root_snapshots(
        paths.root_workbooks,
        paths.extra_input_workbooks,
        paths.a1_outputs,
        paths.generated_extra_inputs,
        roots,
    )
    print(f"Seeded {len(roots)} root(s) from {paths.root_workbooks}: {list(roots)}")
    for folder in written:
        print(f"    {folder}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
