#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
6_sync_og_to_ts20.py

Final A3 stage: propagate the 20-timeslice / 5-bracket fabric from
SOASIA_OSeMOSYS_WV.xlsx down into:
  * OG_csvs_inputs/*.csv  (the source-of-truth A1 reads on the next run)
  * Config_MOMF_T1_A.yaml  (xtra_scen.DailyTimeBracket and xtra_scen.Timeslices)

Without this step, A1 keeps writing 12-ts CSVs and a 12-ts YAML, while A3
already produces 20-ts xlsx outputs. B1_Compiler then aborts with
'These variables are differents...'.

Files rewritten in --og-csvs-dir:
    TIMESLICE.csv               20 entries, S1D1..S4D5
    DAILYTIMEBRACKET.csv        [1..5]
    YearSplit.csv               WV.Yearsplit_Template, broadcast to YEARS
    DaySplit.csv                WV.DaySplit x365 (frac-of-day), broadcast to YEARS
    SpecifiedDemandProfile.csv  WV.Demand_Profiles values (ELC*03 -> ELC*02,
                                REGION='GLOBAL'), broadcast to YEARS
    Conversionls.csv            S?D? -> S? identity (20 ts x 4 seasons)
    Conversionld.csv            always 1 (20 ts x 1 daytype)
    Conversionlh.csv            S?D? -> D? identity (20 ts x 5 brackets)

NOT touched:
    CapacityFactor.csv          per-tech CF roster differs between OG (37 techs)
                                and WV (105 techs); 5 OG-only PWRCSP* techs would
                                be lost on a wholesale rewrite. Handled separately.

Mapping note (SpecifiedDemandProfile):
    WV's Demand_Profiles uses ELC<region>03 codes (post-DSPTRN). The OG CSV is
    pre-DSPTRN (ELC<region>02); A1 + DSPTRN re-apply the 02->03 conversion at
    write-time to A-O_Demand.xlsx. So this script strips '03' back to '02' so
    the OG CSV stays canonical.
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

import pandas as pd

YEARS = list(range(2023, 2051))  # 2023..2050 inclusive
SEASONS = [1, 2, 3, 4]
BRACKETS = [1, 2, 3, 4, 5]
TIMESLICES = [f"S{s}D{d}" for s in SEASONS for d in BRACKETS]  # 20 entries
DAYSPLIT_UNIT_FACTOR = 365.0  # WV stores fraction-of-year; OG CSV expects fraction-of-day


# ---------------------------------------------------------------------------
# CSV writers
# ---------------------------------------------------------------------------
def write_timeslice_csv(out_dir: Path) -> None:
    df = pd.DataFrame({"VALUE": TIMESLICES})
    df.to_csv(out_dir / "TIMESLICE.csv", index=False)
    print(f"  TIMESLICE.csv               {len(df)} rows")


def write_dailytimebracket_csv(out_dir: Path) -> None:
    df = pd.DataFrame({"VALUE": BRACKETS})
    df.to_csv(out_dir / "DAILYTIMEBRACKET.csv", index=False)
    print(f"  DAILYTIMEBRACKET.csv        {len(df)} rows")


def write_yearsplit_csv(wv_file: Path, out_dir: Path) -> None:
    src = pd.read_excel(wv_file, sheet_name="Yearsplit_Template")
    if list(src["Timeslices"]) != TIMESLICES:
        sys.exit(f"[ERROR] WV Yearsplit_Template not in canonical S1D1..S4D5 order")
    # WV has per-year columns 2023..2050; in this build they're constant per ts.
    # Use the 2023 column as the per-ts value (same value for every modeled year).
    ts_to_val = dict(zip(src["Timeslices"], src[2023]))
    rows = [(ts, y, ts_to_val[ts]) for ts in TIMESLICES for y in YEARS]
    df = pd.DataFrame(rows, columns=["TIMESLICE", "YEAR", "VALUE"])
    df.to_csv(out_dir / "YearSplit.csv", index=False)
    total = sum(ts_to_val.values())
    print(f"  YearSplit.csv               {len(df)} rows  (sum-per-year = {total:.6f})")


def write_daysplit_csv(wv_file: Path, out_dir: Path) -> None:
    src = pd.read_excel(wv_file, sheet_name="DaySplit")
    if list(src["DAILYTIMEBRACKET"]) != BRACKETS:
        sys.exit(f"[ERROR] WV DaySplit brackets not [1..5]: {list(src['DAILYTIMEBRACKET'])}")
    # WV stores fraction-of-year (sums to 1/365); OG expects fraction-of-day (sums to 1).
    bracket_to_val = {b: float(src.loc[src["DAILYTIMEBRACKET"] == b, 2023].iloc[0]) * DAYSPLIT_UNIT_FACTOR
                      for b in BRACKETS}
    rows = [(b, y, bracket_to_val[b]) for b in BRACKETS for y in YEARS]
    df = pd.DataFrame(rows, columns=["DAILYTIMEBRACKET", "YEAR", "VALUE"])
    df.to_csv(out_dir / "DaySplit.csv", index=False)
    total = sum(bracket_to_val.values())
    print(f"  DaySplit.csv                {len(df)} rows  (sum-per-year = {total:.6f})")


_ELC_03_RE = re.compile(r"^(ELC[A-Z]{5})03$")


def write_specified_demand_profile_csv(wv_file: Path, out_dir: Path, region: str = "GLOBAL") -> None:
    src = pd.read_excel(wv_file, sheet_name="Demand_Profiles")
    # Map ELC*03 -> ELC*02 to keep the OG CSV pre-DSPTRN.
    def to_pre_dsptrn(code: str) -> str:
        m = _ELC_03_RE.match(str(code))
        return m.group(1) + "02" if m else str(code)
    src = src.copy()
    src["FUEL_OG"] = src["Fuel/Tech"].apply(to_pre_dsptrn)
    if not set(src["Timeslices"]) == set(TIMESLICES):
        sys.exit(f"[ERROR] WV Demand_Profiles timeslices != canonical 20-ts set")
    # Use 2023 column; values are constant across years in this WV build.
    rows = []
    for _, r in src.iterrows():
        val = float(r[2023])
        ts = r["Timeslices"]
        fuel = r["FUEL_OG"]
        for y in YEARS:
            rows.append((region, fuel, ts, y, val))
    df = pd.DataFrame(rows, columns=["REGION", "FUEL", "TIMESLICE", "YEAR", "VALUE"])
    # Stable order: REGION, FUEL, TIMESLICE, YEAR
    df = df.sort_values(["REGION", "FUEL", "TIMESLICE", "YEAR"]).reset_index(drop=True)
    df.to_csv(out_dir / "SpecifiedDemandProfile.csv", index=False)
    print(f"  SpecifiedDemandProfile.csv  {len(df)} rows  ({df['FUEL'].nunique()} fuels x {len(TIMESLICES)} ts x {len(YEARS)} yr)")


def write_conversion_csvs(out_dir: Path) -> None:
    # Conversionls: TIMESLICE x SEASON; 1.0 if S<season> matches the ts season prefix
    rows = [(ts, s, 1.0 if ts.startswith(f"S{s}") else 0.0) for ts in TIMESLICES for s in SEASONS]
    pd.DataFrame(rows, columns=["TIMESLICE", "SEASON", "VALUE"]).to_csv(
        out_dir / "Conversionls.csv", index=False)
    print(f"  Conversionls.csv            {len(rows)} rows")

    # Conversionld: TIMESLICE x DAYTYPE; only one daytype, value=1
    rows = [(ts, 1, 1) for ts in TIMESLICES]
    pd.DataFrame(rows, columns=["TIMESLICE", "DAYTYPE", "VALUE"]).to_csv(
        out_dir / "Conversionld.csv", index=False)
    print(f"  Conversionld.csv            {len(rows)} rows")

    # Conversionlh: TIMESLICE x DAILYTIMEBRACKET; 1.0 if D<bracket> matches the ts bracket suffix
    rows = [(ts, b, 1.0 if ts.endswith(f"D{b}") else 0.0) for ts in TIMESLICES for b in BRACKETS]
    pd.DataFrame(rows, columns=["TIMESLICE", "DAILYTIMEBRACKET", "VALUE"]).to_csv(
        out_dir / "Conversionlh.csv", index=False)
    print(f"  Conversionlh.csv            {len(rows)} rows")


# ---------------------------------------------------------------------------
# YAML editor (mirrors A1's update_yaml_xtra_scen line-replace approach so we
# preserve formatting and comments)
# ---------------------------------------------------------------------------
def update_yaml_xtra_scen(yaml_path: Path) -> None:
    with open(yaml_path, "r", encoding="utf-8") as f:
        lines = f.readlines()

    replacements = {
        "DailyTimeBracket": [f"'{b}'" for b in BRACKETS],
        "Timeslices":       [f"'{ts}'" for ts in TIMESLICES],
    }

    out = []
    i = 0
    while i < len(lines):
        line = lines[i]
        matched = False
        for key, new_values in replacements.items():
            inline_re = rf"^(\s*{key}:\s*)\[.*\](.*)$"
            inline_m = re.match(inline_re, line)
            if inline_m:
                prefix, suffix = inline_m.groups()
                out.append(f"{prefix}[{', '.join(new_values)}]{suffix}\n")
                i += 1
                matched = True
                break
            multi_re = rf"^(\s*){key}:\s*$"
            multi_m = re.match(multi_re, line)
            if multi_m:
                indent = multi_m.group(1)
                out.append(line)
                i += 1
                while i < len(lines) and re.match(rf"^{indent}- ", lines[i]):
                    i += 1
                for v in new_values:
                    out.append(f"{indent}- {v}\n")
                matched = True
                break
        if not matched:
            out.append(line)
            i += 1

    with open(yaml_path, "w", encoding="utf-8") as f:
        f.writelines(out)
    print(f"  YAML xtra_scen.DailyTimeBracket -> [{', '.join(replacements['DailyTimeBracket'])}]")
    print(f"  YAML xtra_scen.Timeslices       -> {len(replacements['Timeslices'])} entries (S1D1..S4D5)")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--wv", type=Path, required=True,
                    help="Path to SOASIA_OSeMOSYS_WV.xlsx (built by A3_process stage1 script 1)")
    ap.add_argument("--og-csvs-dir", type=Path, required=True,
                    help="Path to OG_csvs_inputs/ folder")
    ap.add_argument("--yaml", type=Path, required=True,
                    help="Path to Config_MOMF_T1_A.yaml")
    args = ap.parse_args()

    if not args.wv.is_file():
        sys.exit(f"[ERROR] WV file not found: {args.wv}")
    if not args.og_csvs_dir.is_dir():
        sys.exit(f"[ERROR] OG csvs dir not found: {args.og_csvs_dir}")
    if not args.yaml.is_file():
        sys.exit(f"[ERROR] YAML not found: {args.yaml}")

    print(f"[INFO] WV source        : {args.wv}")
    print(f"[INFO] OG csvs target   : {args.og_csvs_dir}")
    print(f"[INFO] YAML target      : {args.yaml}")
    print(f"[INFO] Years            : {YEARS[0]}..{YEARS[-1]} ({len(YEARS)} years)")
    print(f"[INFO] Timeslices       : {len(TIMESLICES)} (S1D1..S4D5)")
    print()

    print("[INFO] Rewriting OG_csvs_inputs CSVs:")
    write_timeslice_csv(args.og_csvs_dir)
    write_dailytimebracket_csv(args.og_csvs_dir)
    write_yearsplit_csv(args.wv, args.og_csvs_dir)
    write_daysplit_csv(args.wv, args.og_csvs_dir)
    write_specified_demand_profile_csv(args.wv, args.og_csvs_dir)
    write_conversion_csvs(args.og_csvs_dir)
    print()

    print("[INFO] Updating YAML xtra_scen:")
    update_yaml_xtra_scen(args.yaml)
    print()

    print("[INFO] Done. CapacityFactor.csv intentionally NOT touched (per-tech roster mismatch).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
