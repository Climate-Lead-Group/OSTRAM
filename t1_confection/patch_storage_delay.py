#! python3
"""
patch_storage_delay.py
======================

Non-destructive storage-delay patcher for OSTRAM preprocessed datafiles.

This keeps storage in the model, but blocks storage builds for the first N
model years. After those early years, linked storage technologies are opened
again so the optimizer can add storage.

What it writes:
  1. A sibling datafile with:
     - StorageBuildAllowed = 0 for targeted storages in blocked years.
     - PWRLDS*/PWRSDS* TotalAnnualMaxCapacity and
       TotalAnnualMaxCapacityInvestment set to 0 in blocked years.
     - Those same PWR storage caps set to the configured open value
       (default -1, i.e. unconstrained) in later years.
     - Blocking-year minimum capacity/investment rows for those storage
       technologies dropped to avoid min-vs-max conflicts.
  2. A sibling model file with one new parameter and one new constraint:
     NewStorageCapacity[r,s,y] = 0 where StorageBuildAllowed[r,s,y] = 0.

The source datafile and source model are never modified.
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path


STORAGE_ALLOWED_PARAM = "StorageBuildAllowed"
STORAGE_DELAY_CONSTRAINT = "SD1_StorageBuildDelay"

LINK_PARAMS = ("TechnologyToStorage", "TechnologyFromStorage")
TECH_CAP_PARAMS = (
    "TotalAnnualMaxCapacity",
    "TotalAnnualMaxCapacityInvestment",
)
TECH_MIN_PARAMS = (
    "TotalAnnualMinCapacity",
    "TotalAnnualMinCapacityInvestment",
)


def read_lines(path: Path) -> list[str]:
    with open(path, "rb") as f:
        return f.read().decode("utf-8").splitlines(keepends=True)


def write_lines(path: Path, lines: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "wb") as f:
        f.write("".join(lines).encode("utf-8"))


def detect_eol(lines: list[str]) -> str:
    for line in lines:
        if line.endswith("\r\n"):
            return "\r\n"
        if line.endswith("\n"):
            return "\n"
    return "\n"


def find_set_members(lines: list[str], set_name: str) -> list[str]:
    header_re = re.compile(rf"^\s*set\s+{re.escape(set_name)}\b")
    for i, line in enumerate(lines):
        if not header_re.match(line):
            continue
        members: list[str] = []
        j = i + 1
        while j < len(lines) and not lines[j].lstrip().startswith(";"):
            value = lines[j].strip()
            if value:
                members.append(value)
            j += 1
        if j >= len(lines):
            raise RuntimeError(f"Unterminated set {set_name!r}")
        return members
    raise RuntimeError(f"set {set_name!r} not found")


def find_param_block(lines: list[str], param_name: str) -> tuple[int | None, int | None]:
    header_re = re.compile(
        rf"^\s*param(?:\s+default\s+\S+)?\s*:\s*{re.escape(param_name)}\s*:=\s*$"
    )
    for i, line in enumerate(lines):
        if not header_re.match(line):
            continue
        j = i + 1
        while j < len(lines) and not lines[j].lstrip().startswith(";"):
            j += 1
        if j >= len(lines):
            raise RuntimeError(f"Unterminated param block {param_name!r}")
        return i, j
    return None, None


def find_end_line(lines: list[str]) -> int:
    for i, line in enumerate(lines):
        if line.strip() == "end;":
            return i
    return len(lines)


def year_key(year: str) -> tuple[int, str]:
    try:
        return int(year), year
    except ValueError:
        return sys.maxsize, year


def select_blocked_years(years: list[str], first_n_years: int) -> list[str]:
    if first_n_years < 0:
        raise ValueError("--first-n-years must be >= 0")
    sorted_years = sorted(years, key=year_key)
    return sorted_years[:first_n_years]


def select_target_storages(
    storages: list[str],
    exact_storages: list[str],
    storage_prefixes: list[str],
) -> list[str]:
    if exact_storages:
        unknown = sorted(set(exact_storages) - set(storages))
        if unknown:
            raise ValueError(f"Unknown storage names: {unknown}")
        return sorted(exact_storages)
    if storage_prefixes:
        prefixes = tuple(storage_prefixes)
        selected = sorted(s for s in storages if s.startswith(prefixes))
        if not selected:
            raise ValueError(f"No storages matched prefixes: {storage_prefixes}")
        return selected
    return sorted(storages)


def storage_to_tech(storage_name: str) -> str:
    if len(storage_name) < 3:
        raise ValueError(f"Unexpected storage name: {storage_name!r}")
    return "PWR" + storage_name[:-2]


def is_nonzero(value: str) -> bool:
    try:
        return float(value) != 0.0
    except ValueError:
        return bool(value.strip())


def linked_storage_techs(lines: list[str], target_storages: set[str]) -> list[str]:
    techs: set[str] = set()
    for param_name in LINK_PARAMS:
        start, end = find_param_block(lines, param_name)
        if start is None or end is None:
            continue
        for line in lines[start + 1 : end]:
            parts = line.split()
            # REGION TECHNOLOGY STORAGE MODE VALUE
            if len(parts) >= 5 and parts[2] in target_storages and is_nonzero(parts[4]):
                techs.add(parts[1])

    # Naming convention fallback covers sparse link tables and keeps the patch
    # useful on partially stripped/debug datafiles.
    for storage in target_storages:
        techs.add(storage_to_tech(storage))
    return sorted(techs)


def storage_allowed_data_block(
    regions: list[str],
    storages: list[str],
    blocked_years: list[str],
    eol: str,
) -> list[str]:
    block = [f"param {STORAGE_ALLOWED_PARAM} :={eol}"]
    for region in sorted(regions):
        for storage in sorted(storages):
            for year in blocked_years:
                block.append(f"{region} {storage} {year} 0{eol}")
    block.append(f";{eol}")
    return block


def upsert_param_block(lines: list[str], param_name: str, new_block: list[str]) -> list[str]:
    start, end = find_param_block(lines, param_name)
    if start is not None and end is not None:
        return lines[:start] + new_block + lines[end + 1 :]

    insert_at = find_end_line(lines)
    prefix = lines[:insert_at]
    suffix = lines[insert_at:]
    if prefix and prefix[-1].strip():
        prefix = prefix + [detect_eol(lines)]
    return prefix + new_block + suffix


def rewrite_tech_cap_block(
    lines: list[str],
    param_name: str,
    regions: list[str],
    techs: list[str],
    years: list[str],
    blocked_years: set[str],
    allowed_value: str,
    eol: str,
) -> tuple[list[str], int, int]:
    start, end = find_param_block(lines, param_name)
    if start is None or end is None:
        raise RuntimeError(f"param block {param_name!r} not found")

    tech_set = set(techs)
    year_set = set(years)
    removed = 0
    kept: list[str] = []
    for line in lines[start + 1 : end]:
        parts = line.split()
        # REGION TECHNOLOGY YEAR VALUE
        if len(parts) >= 4 and parts[1] in tech_set and parts[2] in year_set:
            removed += 1
            continue
        kept.append(line)

    injected: list[str] = []
    for region in sorted(regions):
        for tech in sorted(techs):
            for year in sorted(years, key=year_key):
                value = "0" if year in blocked_years else str(allowed_value)
                injected.append(f"{region} {tech} {year} {value}{eol}")

    new_lines = lines[: start + 1] + kept + injected + lines[end:]
    return new_lines, removed, len(injected)


def drop_blocked_min_rows(
    lines: list[str],
    param_name: str,
    techs: list[str],
    blocked_years: set[str],
) -> tuple[list[str], int]:
    start, end = find_param_block(lines, param_name)
    if start is None or end is None:
        return lines, 0

    tech_set = set(techs)
    dropped = 0
    kept: list[str] = []
    for line in lines[start + 1 : end]:
        parts = line.split()
        # REGION TECHNOLOGY YEAR VALUE
        if len(parts) >= 4 and parts[1] in tech_set and parts[2] in blocked_years:
            dropped += 1
            continue
        kept.append(line)

    return lines[: start + 1] + kept + lines[end:], dropped


def patch_model_lines(lines: list[str]) -> list[str]:
    eol = detect_eol(lines)
    text = "".join(lines)
    out = list(lines)

    if f"param {STORAGE_ALLOWED_PARAM}" not in text:
        insert_at = None
        for i, line in enumerate(out):
            if "param ResidualStorageCapacity" in line:
                insert_at = i + 1
                break
        if insert_at is None:
            raise RuntimeError("Could not find ResidualStorageCapacity declaration in model")
        out[insert_at:insert_at] = [
            f"param {STORAGE_ALLOWED_PARAM}{{r in REGION, s in STORAGE, y in YEAR}} default 1;{eol}"
        ]

    text = "".join(out)
    if STORAGE_DELAY_CONSTRAINT not in text:
        insert_at = None
        for i, line in enumerate(out):
            if "s.t. SI3_TotalNewStorage" in line:
                insert_at = i + 1
                break
        if insert_at is None:
            raise RuntimeError("Could not find SI3_TotalNewStorage constraint in model")
        out[insert_at:insert_at] = [
            f"s.t. {STORAGE_DELAY_CONSTRAINT}{{r in REGION, s in STORAGE, y in YEAR: {STORAGE_ALLOWED_PARAM}[r,s,y] = 0}}:{eol}",
            f"    NewStorageCapacity[r,s,y] = 0;{eol}",
        ]

    return out


def patch_data_lines(
    lines: list[str],
    first_n_years: int,
    exact_storages: list[str],
    storage_prefixes: list[str],
    allowed_value: str,
) -> tuple[list[str], dict[str, object]]:
    eol = detect_eol(lines)
    regions = find_set_members(lines, "REGION")
    years = sorted(find_set_members(lines, "YEAR"), key=year_key)
    storages = find_set_members(lines, "STORAGE")
    blocked_years = select_blocked_years(years, first_n_years)
    target_storages = select_target_storages(storages, exact_storages, storage_prefixes)
    target_techs = linked_storage_techs(lines, set(target_storages))

    out = upsert_param_block(
        lines,
        STORAGE_ALLOWED_PARAM,
        storage_allowed_data_block(regions, target_storages, blocked_years, eol),
    )

    cap_stats: dict[str, dict[str, int]] = {}
    for param_name in TECH_CAP_PARAMS:
        out, removed, injected = rewrite_tech_cap_block(
            out,
            param_name,
            regions,
            target_techs,
            years,
            set(blocked_years),
            allowed_value,
            eol,
        )
        cap_stats[param_name] = {"removed": removed, "injected": injected}

    min_stats: dict[str, int] = {}
    for param_name in TECH_MIN_PARAMS:
        out, dropped = drop_blocked_min_rows(
            out,
            param_name,
            target_techs,
            set(blocked_years),
        )
        min_stats[param_name] = dropped

    summary = {
        "regions": regions,
        "years": years,
        "blocked_years": blocked_years,
        "target_storages": target_storages,
        "target_techs": target_techs,
        "allowed_value_after_block": allowed_value,
        "cap_stats": cap_stats,
        "min_rows_dropped": min_stats,
    }
    return out, summary


def main() -> int:
    parser = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("input", help="Input preprocessed .txt datafile")
    parser.add_argument("-o", "--output", required=True, help="Output patched datafile")
    parser.add_argument("--model-input", required=True, help="Input OSeMOSYS model file")
    parser.add_argument("--model-output", required=True, help="Output patched model file")
    parser.add_argument(
        "--first-n-years",
        type=int,
        default=5,
        help="Number of earliest model years where storage builds are blocked",
    )
    parser.add_argument(
        "--storage-prefixes",
        nargs="+",
        default=[],
        help="Storage facility prefixes to target, e.g. SDS LDS. Defaults to all storage.",
    )
    parser.add_argument(
        "--storages",
        nargs="+",
        default=[],
        help="Exact storage facility names to target. Overrides --storage-prefixes.",
    )
    parser.add_argument(
        "--allowed-value",
        default="-1",
        help="Cap value for storage-linked PWR techs after blocked years. Default -1 means unconstrained.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Parse and report the patch without writing output files.",
    )
    args = parser.parse_args()

    input_path = Path(args.input)
    output_path = Path(args.output)
    model_input = Path(args.model_input)
    model_output = Path(args.model_output)

    if not input_path.exists():
        sys.exit(f"Input datafile not found: {input_path}")
    if not model_input.exists():
        sys.exit(f"Input model file not found: {model_input}")

    data_lines = read_lines(input_path)
    model_lines = read_lines(model_input)
    patched_data, summary = patch_data_lines(
        data_lines,
        first_n_years=args.first_n_years,
        exact_storages=args.storages,
        storage_prefixes=args.storage_prefixes,
        allowed_value=str(args.allowed_value),
    )
    patched_model = patch_model_lines(model_lines)

    print("=" * 72)
    print("patch_storage_delay.py")
    print(f"  Input data:      {input_path}")
    print(f"  Output data:     {output_path}")
    print(f"  Input model:     {model_input}")
    print(f"  Output model:    {model_output}")
    print(f"  First N years:   {args.first_n_years}")
    print(f"  Blocked years:   {summary['blocked_years']}")
    print(f"  Target storages: {len(summary['target_storages'])}")
    print(f"  Target techs:    {len(summary['target_techs'])}")
    print(f"  Later cap value: {summary['allowed_value_after_block']}")
    for param_name, stats in summary["cap_stats"].items():
        print(
            f"  {param_name}: removed {stats['removed']} existing rows, "
            f"injected {stats['injected']} rows"
        )
    for param_name, dropped in summary["min_rows_dropped"].items():
        print(f"  {param_name}: dropped {dropped} blocked-year minimum rows")

    if args.dry_run:
        print("  dry-run: no files written")
        return 0

    write_lines(output_path, patched_data)
    write_lines(model_output, patched_model)
    print(f"  wrote data:  {output_path} ({output_path.stat().st_size:,} bytes)")
    print(f"  wrote model: {model_output} ({model_output.stat().st_size:,} bytes)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
