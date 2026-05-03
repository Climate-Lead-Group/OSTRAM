#!/usr/bin/env python3
"""
test_strip_storage.py - Run strip_storage.py + validate the output in one shot.

Takes the same args as strip_storage.py. Runs the strip as a subprocess, then
walks the output file structurally (set blocks + parameter blocks) and verifies:

    1. STORAGE set excludes the disabled facilities.
    2. Every STORAGE_PARAMS block has zero rows referencing disabled storages.
    3. Every LINK_PARAMS block (col 3 = STORAGE) has zero references.
    4. Every MODExTECHNOLOGYperSTORAGE[to|from][<disabled>] aux set is gone.
    5. TotalAnnualMaxCapacity has exactly (n_years x n_disabled_techs) rows for
       the disabled techs, all with value = 0.
    6. Every TECH_MIN_PARAMS block has zero rows for the disabled techs
       (this is the NCC2 contradiction guard).
    7. No bare PWRSDS / PWRLDS phantom tech anywhere (Codex regression check).
    8. Whole-file scan: zero residual mentions of any disabled storage facility.

Exit code 0 if all pass, 1 if any fail.

Usage (same args as strip_storage.py, plus --strip-script if it lives elsewhere):

    python test_strip_storage.py Pre_processed_BAU_0.txt \
        -o Pre_processed_BAU_0_NoLKASDS.txt \
        --mode tech --targets SDSLKAXX01

    python test_strip_storage.py Pre_processed_BAU_0.txt \
        -o Pre_processed_BAU_0_NoStorage.txt \
        --mode all
"""

from __future__ import annotations
import argparse
import re
import subprocess
import sys
from pathlib import Path


def count_param_block_rows(lines, param_name, key_col_1indexed, key_set):
    """Count rows in `param ... : <param_name> := ... ;` block where
    tokens[key_col-1] is in key_set."""
    header_re = re.compile(rf"^\s*param.*:\s*{re.escape(param_name)}\s*:=\s*$")
    count = 0
    in_block = False
    for line in lines:
        if header_re.match(line):
            in_block = True
            continue
        if in_block:
            if line.lstrip().startswith(";"):
                in_block = False
                continue
            tokens = line.split()
            if len(tokens) >= key_col_1indexed and tokens[key_col_1indexed - 1] in key_set:
                count += 1
    return count


def collect_param_block_rows(lines, param_name, key_col_1indexed, key_set):
    """Return matching row strings (for diagnostics)."""
    header_re = re.compile(rf"^\s*param.*:\s*{re.escape(param_name)}\s*:=\s*$")
    out = []
    in_block = False
    for line in lines:
        if header_re.match(line):
            in_block = True
            continue
        if in_block:
            if line.lstrip().startswith(";"):
                in_block = False
                continue
            tokens = line.split()
            if len(tokens) >= key_col_1indexed and tokens[key_col_1indexed - 1] in key_set:
                out.append(line.rstrip())
    return out


def main():
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("input")
    ap.add_argument("-o", "--output", required=True)
    ap.add_argument("--mode", choices=["tech", "class", "all"], required=True)
    ap.add_argument("--targets", nargs="+", default=[])
    ap.add_argument("--strip-script", default="strip_storage.py",
                    help="Path to strip_storage.py (default: ./strip_storage.py)")
    args = ap.parse_args()

    in_path = Path(args.input)
    out_path = Path(args.output)
    strip_path = Path(args.strip_script).resolve()

    if not in_path.exists():
        sys.exit(f"Input not found: {in_path}")
    if not strip_path.exists():
        sys.exit(f"strip_storage.py not found at: {strip_path}")

    # Import config + helpers from the strip script we're testing
    sys.path.insert(0, str(strip_path.parent))
    import strip_storage as ss

    # Read input first so we know the YEAR set + initial STORAGE membership
    in_lines = ss.read_lines(in_path)
    years = ss.find_year_set(in_lines)
    _, _, in_storages = ss.find_storage_set(in_lines)

    # --- Run strip_storage.py as a subprocess ---
    cmd = [sys.executable, str(strip_path), str(in_path),
           "-o", str(out_path), "--mode", args.mode]
    if args.targets:
        cmd += ["--targets", *args.targets]

    print("=" * 72)
    print("STEP 1: Running strip_storage.py")
    print("=" * 72)
    print("CMD:", " ".join(cmd))
    print()
    res = subprocess.run(cmd, capture_output=True, text=True)
    print(res.stdout, end="")
    if res.returncode != 0:
        print("STDERR:", res.stderr)
        sys.exit(f"strip_storage exited with code {res.returncode}")
    print()

    # --- Validate output ---
    out_lines = ss.read_lines(out_path)
    _, _, out_storages = ss.find_storage_set(out_lines)
    disabled_storages = sorted(set(in_storages) - set(out_storages))
    disabled_techs = sorted({ss.storage_to_tech(s) for s in disabled_storages})

    print("=" * 72)
    print("STEP 2: Validating output")
    print("=" * 72)
    print(f"Disabled storages ({len(disabled_storages)}): {disabled_storages}")
    print(f"Disabled techs    ({len(disabled_techs)}): {disabled_techs}")
    print(f"YEAR set size:     {len(years)}")
    print()

    failures = []

    def check(name, expected, actual, detail=None):
        ok = (expected == actual)
        status = "PASS" if ok else "FAIL"
        msg = f"  [{status}] {name}: expected={expected} actual={actual}"
        if detail and not ok:
            msg += f"\n         {detail}"
        print(msg)
        if not ok:
            failures.append(name)
        return ok

    # 1. STORAGE set excludes disabled facilities
    print("(1) STORAGE set membership")
    check("disabled facilities removed from `set STORAGE`",
          0, sum(1 for s in disabled_storages if s in out_storages))

    # 2. STORAGE_PARAMS (col 2 = STORAGE)
    print("\n(2) Storage-side parameter blocks (col 2 = STORAGE)")
    for p in ss.STORAGE_PARAMS:
        cnt = count_param_block_rows(out_lines, p, 2, set(disabled_storages))
        check(f"  {p}: residual rows for disabled storages", 0, cnt)

    # 3. LINK_PARAMS (col 3 = STORAGE)
    print("\n(3) Tech<->storage link params (col 3 = STORAGE)")
    for p in ss.LINK_PARAMS:
        cnt = count_param_block_rows(out_lines, p, 3, set(disabled_storages))
        check(f"  {p}: residual rows for disabled storages", 0, cnt)

    # 4. Aux derived sets removed
    print("\n(4) Auxiliary MODExTECHNOLOGYperSTORAGE[to|from] sets")
    aux_pat = re.compile(r"set\s+MODExTECHNOLOGYperSTORAGE(?:to|from)\[([^\]]+)\]")
    leaked_aux = []
    for line in out_lines:
        m = aux_pat.search(line)
        if m and m.group(1) in disabled_storages:
            leaked_aux.append(line.rstrip())
    check("aux sets for disabled storages removed",
          0, len(leaked_aux),
          detail=f"sample leaks: {leaked_aux[:3]}")

    # 5. TotalAnnualMaxCapacity = 0 injected for disabled techs
    print("\n(5) TotalAnnualMaxCapacity zero-row injection")
    expected_max_rows = len(years) * len(disabled_techs)
    actual_max_rows = count_param_block_rows(
        out_lines, "TotalAnnualMaxCapacity", 2, set(disabled_techs))
    check("zero-row count = n_years x n_disabled_techs",
          expected_max_rows, actual_max_rows)

    # check those rows are actually zero (col 4 = VALUE)
    bad_values = []
    header_re = re.compile(r"^\s*param.*:\s*TotalAnnualMaxCapacity\s*:=\s*$")
    in_block = False
    for line in out_lines:
        if header_re.match(line):
            in_block = True
            continue
        if in_block:
            if line.lstrip().startswith(";"):
                break
            tokens = line.split()
            if len(tokens) >= 4 and tokens[1] in disabled_techs:
                if tokens[3] != "0":
                    bad_values.append(line.rstrip())
    check("all injected values exactly 0",
          0, len(bad_values),
          detail=f"non-zero rows: {bad_values[:3]}")

    # 6. TECH_MIN_PARAMS (col 2 = TECH) — the NCC2 contradiction guard
    print("\n(6) Tech-side min parameter blocks (col 2 = TECH) — NCC2 guard")
    for p in ss.TECH_MIN_PARAMS:
        rows = collect_param_block_rows(out_lines, p, 2, set(disabled_techs))
        check(f"  {p}: residual rows for disabled techs",
              0, len(rows),
              detail=f"leaked rows: {rows[:3]}")

    # 7. Codex regression: bare PWRSDS / PWRLDS (without country suffix)
    print("\n(7) Codex regression check")
    bare_re = re.compile(r"\bPWR(?:SDS|LDS)\b(?!\w)")
    bare_hits = []
    for ln, line in enumerate(out_lines, 1):
        if bare_re.search(line):
            bare_hits.append(f"line {ln}: {line.rstrip()}")
    check("no bare PWRSDS/PWRLDS phantom tech",
          0, len(bare_hits),
          detail=f"hits: {bare_hits[:3]}")

    # 8. Whole-file scan for any residual reference to disabled storages
    print("\n(8) Whole-file residual mentions of disabled storage facilities")
    for s in disabled_storages:
        pat = re.compile(rf"(?<!\w){re.escape(s)}(?!\w)")
        hits = sum(1 for line in out_lines if pat.search(line))
        check(f"  {s}: residual mentions", 0, hits)

    print()
    print("=" * 72)
    if failures:
        print(f"OVERALL: FAIL ({len(failures)} check(s) failed)")
        for f in failures:
            print(f"  - {f}")
        print("=" * 72)
        sys.exit(1)
    else:
        print("OVERALL: PASS — output is ready for B2.py / glpsol")
        print("=" * 72)
        sys.exit(0)


if __name__ == "__main__":
    main()
