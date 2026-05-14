#!/usr/bin/env python3
"""
strip_storage.py - Diagnostic patcher for OSTRAM Pre_processed_*.txt files.

Produces a NEW .txt file with selected storage facilities + their feeding
PWR technologies disabled. The source file is never modified.

Designed to plug into the existing B2 pipeline pattern (cf. inject_DaysInDayType.py)
but invoked manually as a one-off diagnostic step between the DaysInDayType
injector and the `glpsol --wlp` LP-build step.

Three modes:
  --mode tech   --targets SDSLKAXX01 [SDSBGDXX01 ...]
        Disable specific storage facilities by exact name.
  --mode class  --targets SDS [LDS]
        Disable all facilities whose name starts with the given prefix.
  --mode all
        Disable every storage facility in the model.

What "disable" does, atomically:
  1. Removes the facility from `set STORAGE :=`.
  2. Strips rows for the facility from every storage-side parameter block:
       CapitalCostStorage, DiscountRateStorage, MinStorageCharge,
       OperationalLifeStorage, ResidualStorageCapacity, StorageLevelStart,
       StorageMaxChargeRate, StorageMaxDischargeRate.
  3. Strips rows from TechnologyToStorage and TechnologyFromStorage that
     reference the disabled storage.
  4. Removes the auxiliary derived sets `MODExTECHNOLOGYperSTORAGEto[X]`
     and `MODExTECHNOLOGYperSTORAGEfrom[X]` for X in the disabled list.
  5. Injects TotalAnnualMaxCapacity = 0 for every year for the
     corresponding PWR storage-using technology (PWRSDSLKAXX for SDSLKAXX01,
     etc.) so that even if the tech is otherwise referenced, no capacity
     can be built. Existing TotalAnnualMaxCapacity rows for these techs
     are dropped first to avoid duplicate-tuple errors.

Usage:
    python strip_storage.py Pre_processed_BAU_0.txt \
        -o Pre_processed_BAU_0_NoStorage.txt \
        --mode all

    python strip_storage.py Pre_processed_BAU_0.txt \
        -o Pre_processed_BAU_0_NoLKASDS.txt \
        --mode tech --targets SDSLKAXX01

    python strip_storage.py Pre_processed_BAU_0.txt \
        -o Pre_processed_BAU_0_NoSDS.txt \
        --mode class --targets SDS

After patching, run the rest of the pipeline manually:
    glpsol -m <osemosys_model.txt> -d <patched.txt> \
        --wlp <patched>.lp --check
    cplex -c "read <patched>.lp" "optimize" "write <patched>.sol"
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

# -----------------------------------------------------------------------------
# Configuration
# -----------------------------------------------------------------------------

STORAGE_PARAMS = [
    "CapitalCostStorage",
    "DiscountRateStorage",
    "MinStorageCharge",
    "OperationalLifeStorage",
    "ResidualStorageCapacity",
    "StorageLevelStart",
    "StorageMaxChargeRate",
    "StorageMaxDischargeRate",
]

# Parameters where col 3 (1-indexed) is the STORAGE name:
#   GLOBAL <TECH> <STORAGE> <MODE> <VALUE>
LINK_PARAMS = ["TechnologyToStorage", "TechnologyFromStorage"]

# Tech-side "minimum" parameters that can force a disabled PWR tech to be
# built or to operate. Must be stripped for a clean disable; if any of these
# have rows for our targeted techs, a TotalAnnualMaxCapacity=0 alone creates
# a direct min-vs-max contradiction (presolve infeasibility).
# Keyed (REGION, TECHNOLOGY, YEAR, VALUE) for the first three; tech is at col 2.
# TotalTechnologyModelPeriodActivityLowerLimit is keyed (REGION, TECHNOLOGY, VALUE)
# but tech is also at col 2, so the same drop logic applies.
TECH_MIN_PARAMS = [
    "TotalAnnualMinCapacity",
    "TotalAnnualMinCapacityInvestment",
    "TotalTechnologyAnnualActivityLowerLimit",
    "TotalTechnologyModelPeriodActivityLowerLimit",
]

LINE_END = "\r\n"  # Source file uses Windows line endings; preserve them.


# -----------------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------------

def storage_to_tech(storage_name: str) -> str:
    """Map a storage facility name to its driving PWR technology.

    Convention in this model:  SDSLKAXX01 -> PWRSDSLKAXX
                               LDSBGDXX01 -> PWRLDSBGDXX
    Drop the trailing two-character suffix (e.g. '01') and prepend 'PWR'.
    """
    if len(storage_name) < 3:
        raise ValueError(f"Unexpected storage name: {storage_name!r}")
    return "PWR" + storage_name[:-2]


def expand_targets(all_storages: list[str], mode: str, targets: list[str]) -> set[str]:
    """Resolve user-specified targets into the full set of facilities to strip."""
    if mode == "all":
        return set(all_storages)
    if mode == "tech":
        unknown = [t for t in targets if t not in all_storages]
        if unknown:
            raise ValueError(
                f"Unknown storage facility names: {unknown}. "
                f"Valid names: {sorted(all_storages)}"
            )
        return set(targets)
    if mode == "class":
        prefixes = tuple(targets)
        matched = {s for s in all_storages if s.startswith(prefixes)}
        if not matched:
            raise ValueError(
                f"No storage facility matched class prefixes: {targets}"
            )
        return matched
    raise ValueError(f"Unknown mode: {mode}")


# -----------------------------------------------------------------------------
# File I/O (CRLF-preserving)
# -----------------------------------------------------------------------------

def read_lines(path: Path) -> list[str]:
    """Read file preserving line endings."""
    with open(path, "rb") as f:
        raw = f.read()
    return raw.decode("utf-8").splitlines(keepends=True)


def write_lines(path: Path, lines: list[str]) -> None:
    with open(path, "wb") as f:
        f.write("".join(lines).encode("utf-8"))


# -----------------------------------------------------------------------------
# Set / param block surgery
# -----------------------------------------------------------------------------

def find_storage_set(lines: list[str]) -> tuple[int, int, list[str]]:
    """Locate `set STORAGE :=` block. Returns (start, end, members)
    where start is the header line index and end is the `;` line index."""
    for i, line in enumerate(lines):
        if line.lstrip().startswith("set STORAGE"):
            j = i + 1
            members: list[str] = []
            while j < len(lines) and not lines[j].lstrip().startswith(";"):
                m = lines[j].strip()
                if m:
                    members.append(m)
                j += 1
            if j >= len(lines):
                raise RuntimeError("Unterminated `set STORAGE` block (missing ;)")
            return i, j, members
    raise RuntimeError("`set STORAGE` not found in file.")


def find_year_set(lines: list[str]) -> list[str]:
    """Return the list of YEAR set members."""
    for i, line in enumerate(lines):
        if line.lstrip().startswith("set YEAR"):
            j = i + 1
            years: list[str] = []
            while j < len(lines) and not lines[j].lstrip().startswith(";"):
                y = lines[j].strip()
                if y:
                    years.append(y)
                j += 1
            return years
    raise RuntimeError("`set YEAR` not found in file.")


def strip_set_storage(lines: list[str], to_remove: set[str]) -> list[str]:
    start, end, members = find_storage_set(lines)
    kept = [m for m in members if m not in to_remove]
    new_block = [lines[start]]  # header `set STORAGE :=`
    for m in kept:
        new_block.append(m + LINE_END)
    new_block.append(";" + LINE_END)
    return lines[:start] + new_block + lines[end + 1:]


def strip_param_block_rows(
    lines: list[str],
    param_name: str,
    key_col_1indexed: int,
    drop_if,
) -> list[str]:
    """Remove rows from `param ... : <param_name> := ... ;` block where
    drop_if(value_at_key_col) returns True. Token columns are 1-indexed.
    Header and trailing `;` are preserved."""
    header_re = re.compile(rf"^\s*param.*:\s*{re.escape(param_name)}\s*:=\s*$")
    out: list[str] = []
    i = 0
    while i < len(lines):
        if header_re.match(lines[i]):
            out.append(lines[i])
            i += 1
            while i < len(lines) and not lines[i].lstrip().startswith(";"):
                tokens = lines[i].split()
                if (
                    len(tokens) >= key_col_1indexed
                    and drop_if(tokens[key_col_1indexed - 1])
                ):
                    pass  # drop this row
                else:
                    out.append(lines[i])
                i += 1
            if i < len(lines):
                out.append(lines[i])  # `;`
                i += 1
        else:
            out.append(lines[i])
            i += 1
    return out


def strip_aux_sets(lines: list[str], to_remove: set[str]) -> list[str]:
    """Remove auxiliary derived sets like
        set MODExTECHNOLOGYperSTORAGEto[SDSLKAXX01]:= (1, PWRSDSLKAXX);
    when the bracketed storage is in to_remove."""
    pattern = re.compile(r"set\s+MODExTECHNOLOGYperSTORAGE(?:to|from)\[([^\]]+)\]")
    out: list[str] = []
    for line in lines:
        m = pattern.search(line)
        if m and m.group(1) in to_remove:
            continue
        out.append(line)
    return out


def inject_zero_max_capacity(lines: list[str], techs: set[str]) -> list[str]:
    """Within the `param ... : TotalAnnualMaxCapacity := ... ;` block,
    drop any existing rows for `techs` and replace with explicit zeros
    for every YEAR. This forces the storage-using PWR techs to be
    non-investable."""
    years = find_year_set(lines)
    header_re = re.compile(r"^\s*param.*:\s*TotalAnnualMaxCapacity\s*:=\s*$")

    out: list[str] = []
    i = 0
    handled = False
    while i < len(lines):
        if header_re.match(lines[i]) and not handled:
            out.append(lines[i])
            i += 1
            existing_kept: list[str] = []
            while i < len(lines) and not lines[i].lstrip().startswith(";"):
                tokens = lines[i].split()
                # Row format: GLOBAL <TECH> <YEAR> <VALUE>
                if len(tokens) >= 2 and tokens[1] in techs:
                    pass  # drop
                else:
                    existing_kept.append(lines[i])
                i += 1
            out.extend(existing_kept)
            for t in sorted(techs):
                for y in years:
                    out.append(f"GLOBAL {t} {y} 0{LINE_END}")
            if i < len(lines):
                out.append(lines[i])  # `;`
                i += 1
            handled = True
        else:
            out.append(lines[i])
            i += 1

    if not handled:
        sys.stderr.write(
            "Warning: TotalAnnualMaxCapacity block not found; "
            "tech-side disable was NOT applied.\n"
        )
    return lines if not handled else out


# -----------------------------------------------------------------------------
# Main
# -----------------------------------------------------------------------------

def main() -> int:
    ap = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    ap.add_argument("input", help="Input preprocessed .txt file")
    ap.add_argument("-o", "--output", required=True, help="Output .txt file")
    ap.add_argument(
        "--mode", choices=["tech", "class", "all"], required=True,
        help=(
            "tech: disable specific storage facility names. "
            "class: disable by name prefix (e.g. SDS, LDS). "
            "all: disable every storage."
        ),
    )
    ap.add_argument(
        "--targets", nargs="+", default=[],
        help="For tech mode: facility names. For class mode: prefixes.",
    )
    ap.add_argument(
        "--no-tech-cap-zero", action="store_true",
        help=(
            "Skip injecting TotalAnnualMaxCapacity=0 for the "
            "storage-using PWR techs (NOT recommended)."
        ),
    )
    args = ap.parse_args()

    in_path = Path(args.input)
    out_path = Path(args.output)
    if not in_path.exists():
        sys.exit(f"Input file not found: {in_path}")
    if args.mode != "all" and not args.targets:
        sys.exit(f"--mode {args.mode} requires --targets.")

    lines = read_lines(in_path)
    _, _, all_storages = find_storage_set(lines)

    try:
        to_remove_storages = expand_targets(all_storages, args.mode, args.targets)
    except ValueError as e:
        sys.exit(str(e))

    to_remove_techs = {storage_to_tech(s) for s in to_remove_storages}

    print("=" * 70)
    print(f"strip_storage.py")
    print(f"  Input:  {in_path}")
    print(f"  Output: {out_path}")
    print(f"  Mode:   {args.mode}")
    print(f"  Disabling {len(to_remove_storages)} storage facility/ies:")
    for s in sorted(to_remove_storages):
        print(f"    - {s:<14} (tech: {storage_to_tech(s)})")
    print("=" * 70)

    # 1. Set membership
    lines = strip_set_storage(lines, to_remove_storages)

    # 2. Storage-side parameters: col 2 = STORAGE
    for p in STORAGE_PARAMS:
        lines = strip_param_block_rows(
            lines, p, key_col_1indexed=2,
            drop_if=lambda v, _set=to_remove_storages: v in _set,
        )

    # 3. Tech<->storage links: col 3 = STORAGE
    for p in LINK_PARAMS:
        lines = strip_param_block_rows(
            lines, p, key_col_1indexed=3,
            drop_if=lambda v, _set=to_remove_storages: v in _set,
        )

    # 4. Auxiliary derived sets
    lines = strip_aux_sets(lines, to_remove_storages)

    # 5. Tech-side disable: TotalAnnualMaxCapacity = 0
    if not args.no_tech_cap_zero:
        lines = inject_zero_max_capacity(lines, to_remove_techs)

    # 6. Strip rows from minimum-capacity / minimum-activity parameter blocks
    #    for the disabled techs. Default for all of these is 0, so dropping
    #    rows is equivalent to setting to 0. Without this, a forced minimum
    #    (e.g. TotalAnnualMinCapacityInvestment) collides with the max=0 from
    #    step 5 and the LP is presolve-infeasible.
    for p in TECH_MIN_PARAMS:
        lines = strip_param_block_rows(
            lines, p, key_col_1indexed=2,
            drop_if=lambda v, _set=to_remove_techs: v in _set,
        )

    write_lines(out_path, lines)
    print(f"\nWrote: {out_path}  ({out_path.stat().st_size:,} bytes)")
    print("Next: build the LP and solve, e.g.")
    print(f"  glpsol -m <model.txt> -d {out_path} --wlp {out_path.with_suffix('.lp')} --check")
    print(f"  cplex -c \"read {out_path.with_suffix('.lp')}\" \"optimize\" "
          f"\"write {out_path.with_suffix('.sol')}\"")
    return 0


if __name__ == "__main__":
    sys.exit(main())
