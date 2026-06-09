"""
concat_all_scenarios.py — merge OSTRAM per-scenario Input + output CSVs
=======================================================================
Mirrors the concatenation logic in B2_Executing_OG_Model.py
(`concatenate_all_scenarios`).

WHY THE REWRITE
---------------
The previous version did:

    merged = pd.merge(df_in, df_out, on=common, how='outer')

where `common` was the 11 OSeMOSYS index columns (REGION, YEAR, TECHNOLOGY,
FUEL, EMISSION, MODE_OF_OPERATION, TIMESLICE, STORAGE, SEASON, DAYTYPE,
DAILYTIMEBRACKET). Input parameters and output results live at *different*
index granularities, so most rows are sparse (FUEL/MODE/TIMESLICE = NaN).
Merging on those keys makes every (R,T,Y)-granular OUTPUT row match every
(R,T,Y)-granular INPUT row that shares the same key — so output values get
copied across multiple input-parameter rows (the "outputs repeat" /
inflation symptom).

B2 never merges. It STACKS input rows and output rows vertically with
pd.concat, then aligns columns. Input rows stay input rows, output rows stay
output rows — nothing is duplicated or scattered. This script copies that.

Usage
-----
    python concat_all_scenarios.py
    python concat_all_scenarios.py --search-dir Executables \
        --scenarios BAU,A_Calibrated_BAU,B_Optimised_VRE
"""
import os
import sys
import argparse
from datetime import date

import numpy as np
import pandas as pd

# OSeMOSYS index/set columns, kept at the front of the output
KEYS_SETS = [
    "REGION", "YEAR", "TECHNOLOGY", "FUEL", "EMISSION", "MODE_OF_OPERATION",
    "TIMESLICE", "STORAGE", "SEASON", "DAYTYPE", "DAILYTIMEBRACKET",
]


def reorder_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Metadata + index columns first, everything else alphabetical."""
    front = [c for c in ["Future", "Scenario"] if c in df.columns]
    front += [c for c in KEYS_SETS if c in df.columns]
    rest = sorted(c for c in df.columns if c not in front)
    return df[front + rest]


def sort_rows(df: pd.DataFrame) -> pd.DataFrame:
    sort_cols = [c for c in ["Future", "Scenario", "REGION", "TECHNOLOGY", "YEAR"]
                 if c in df.columns]
    if sort_cols:
        df = df.sort_values(by=sort_cols).reset_index(drop=True)
    return df


def add_accumulated_min_cap_investment(df: pd.DataFrame) -> pd.DataFrame:
    """AccumulatedTotalAnnualMinCapacityInvestment: cumsum over YEAR within
    each (Future, Scenario, TECHNOLOGY) group. Matches B2."""
    if "TotalAnnualMinCapacityInvestment" not in df.columns:
        return df
    df = df.copy()
    df["AccumulatedTotalAnnualMinCapacityInvestment"] = np.nan
    group_cols = [c for c in ["Future", "Scenario", "TECHNOLOGY"] if c in df.columns]
    mask = df["TotalAnnualMinCapacityInvestment"].notna()
    if group_cols:
        df = df.sort_values(by=group_cols + ["YEAR"]).reset_index(drop=True)
        mask = df["TotalAnnualMinCapacityInvestment"].notna()
        df.loc[mask, "AccumulatedTotalAnnualMinCapacityInvestment"] = (
            df.loc[mask]
            .groupby(group_cols, sort=False)["TotalAnnualMinCapacityInvestment"]
            .cumsum()
        )
    else:
        df.loc[mask, "AccumulatedTotalAnnualMinCapacityInvestment"] = (
            df.loc[mask, "TotalAnnualMinCapacityInvestment"].cumsum()
        )
    return df


def find_output_csv(subdir: str, files: list[str]) -> str | None:
    """Prefer the Pre_processed chained-suffix output; fall back to any
    *_output.csv / *_Output.csv."""
    preproc = [f for f in files if f.endswith("_output.csv") and "Pre_processed" in f]
    if preproc:
        # longest name = most chained suffixes = the active solver output
        return os.path.join(subdir, sorted(preproc, key=len)[-1])
    other = [f for f in files if f.lower().endswith("_output.csv")]
    return os.path.join(subdir, sorted(other, key=len)[-1]) if other else None


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--search-dir", default="Executables")
    parser.add_argument("--output", default="OSTRAM_Combined_Inputs_Outputs.csv")
    parser.add_argument("--inputs-file", default="OSTRAM_Combined_Inputs.csv")
    parser.add_argument("--outputs-file", default="OSTRAM_Combined_Outputs.csv")
    parser.add_argument("--scenarios", default=None,
                        help="comma-separated scenario names to keep")
    parser.add_argument("--dated-copies", action="store_true",
                        help="also write _YYYY-MM-DD dated copies")
    args = parser.parse_args()

    if not os.path.isdir(args.search_dir):
        sys.exit(f"[ERROR] Not found: {args.search_dir}")

    scen_filter = ([s.strip() for s in args.scenarios.split(",")]
                   if args.scenarios else None)

    combined_inputs, combined_outputs = [], []

    for entry in sorted(os.listdir(args.search_dir)):
        if entry.lower() in ("default", "__pycache__"):
            continue
        subdir = os.path.join(args.search_dir, entry)
        if not os.path.isdir(subdir):
            continue

        # "A_Calibrated_BAU_0" -> scenario="A_Calibrated_BAU", future="0"
        scenario, _, future = entry.rpartition("_")
        if not scenario:           # no "_N" suffix
            scenario, future = entry, "0"
        if scen_filter and scenario not in scen_filter:
            continue

        files = sorted(os.listdir(subdir))
        input_csv = next((os.path.join(subdir, f) for f in files
                          if f.endswith("_Input.csv")), None)
        output_csv = find_output_csv(subdir, files)

        print(f"[{scenario}  future={future}]")

        if input_csv and os.path.exists(input_csv):
            df_in = pd.read_csv(input_csv, low_memory=False)
            df_in.insert(0, "Future", future)
            df_in.insert(1, "Scenario", scenario)
            combined_inputs.append(df_in)
            print(f"  Input:  {os.path.basename(input_csv)}  ({len(df_in):,} rows)")
        else:
            print("  Input:  MISSING")

        if output_csv and os.path.exists(output_csv):
            df_out = pd.read_csv(output_csv, low_memory=False)
            df_out.insert(0, "Future", future)
            df_out.insert(1, "Scenario", scenario)
            combined_outputs.append(df_out)
            print(f"  Output: {os.path.basename(output_csv)}  ({len(df_out):,} rows)")
        else:
            print("  Output: MISSING")

    if not combined_inputs and not combined_outputs:
        sys.exit("[ERROR] No scenario data found.")

    # ── stack vertically (THE FIX — no merge) ─────────────────────────
    df_inputs_all = (pd.concat(combined_inputs, ignore_index=True)
                     if combined_inputs else pd.DataFrame())
    df_outputs_all = (pd.concat(combined_outputs, ignore_index=True)
                      if combined_outputs else pd.DataFrame())
    df_combined = pd.concat([df_inputs_all, df_outputs_all],
                            ignore_index=True, sort=True)

    today = date.today().isoformat()

    def write(df, path, accumulate=False):
        if df.empty:
            return
        df = sort_rows(reorder_columns(df))
        if accumulate:
            df = add_accumulated_min_cap_investment(df)
            df = sort_rows(df)   # re-sort after the group cumsum reorder
        df.to_csv(path, index=False)
        if args.dated_copies:
            df.to_csv(path.replace(".csv", f"_{today}.csv"), index=False)

    n_in, n_out, n_comb = len(df_inputs_all), len(df_outputs_all), len(df_combined)

    # write the smaller frames first, then free them before the big combined write
    write(df_inputs_all, args.inputs_file)
    write(df_outputs_all, args.outputs_file)
    del df_inputs_all, df_outputs_all
    import gc
    gc.collect()
    write(df_combined, args.output, accumulate=True)

    # ── summary ───────────────────────────────────────────────────────
    print(f"\n{'='*60}")
    print(f"  inputs  : {n_in:,} rows -> {args.inputs_file}")
    print(f"  outputs : {n_out:,} rows -> {args.outputs_file}")
    print(f"  combined: {n_comb:,} rows -> {args.output}")
    print(f"  (= inputs + outputs, no duplication)")
    if not df_combined.empty and "Scenario" in df_combined.columns:
        print(f"{'-'*60}")
        for s in sorted(df_combined["Scenario"].unique()):
            sub = df_combined[df_combined["Scenario"] == s]
            cap = pd.to_numeric(sub.get("TotalCapacityAnnual"), errors="coerce").sum()
            gen = pd.to_numeric(sub.get("ProductionByTechnologyAnnual"), errors="coerce").sum()
            cost = pd.to_numeric(sub.get("TotalDiscountedCost"), errors="coerce").sum()
            print(f"  {s}: {len(sub):,} rows  cap={cap:,.0f}  gen={gen:,.0f}  cost={cost:,.0f}")
    mb = os.path.getsize(args.output) / (1024 * 1024)
    print(f"\n[DONE] {args.output} ({mb:.1f} MB)")


if __name__ == "__main__":
    main()
