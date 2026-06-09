"""
concat_all_scenarios.py — merge OSTRAM per-scenario Input + output CSVs
Usage:  python concat_all_scenarios.py --scenarios BAU,A_Calibrated_BAU,B_Optimised_VRE
"""
import os, sys, argparse
import pandas as pd

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--search-dir', default='Executables')
    parser.add_argument('--output', default='OSTRAM_Combined_Inputs_Outputs.csv')
    parser.add_argument('--scenarios', default=None)
    args = parser.parse_args()

    if not os.path.isdir(args.search_dir):
        sys.exit(f"[ERROR] Not found: {args.search_dir}")

    scen_filter = None
    if args.scenarios:
        scen_filter = [s.strip() for s in args.scenarios.split(',')]

    frames = []
    for entry in sorted(os.listdir(args.search_dir)):
        subdir = os.path.join(args.search_dir, entry)
        if not os.path.isdir(subdir):
            continue
        scen_name = entry.replace('_0', '') if entry.endswith('_0') else entry
        if scen_filter and scen_name not in scen_filter:
            continue

        input_csv = None
        output_csv = None
        for f in sorted(os.listdir(subdir)):
            if f.endswith('_Input.csv'):
                input_csv = os.path.join(subdir, f)
            if f.endswith('_output.csv') and 'Pre_processed' in f:
                output_csv = os.path.join(subdir, f)

        if input_csv and output_csv:
            print(f"[{scen_name}]")
            print(f"  Input:  {os.path.basename(input_csv)}")
            print(f"  Output: {os.path.basename(output_csv)}")
            try:
                df_in = pd.read_csv(input_csv, low_memory=False)
                df_out = pd.read_csv(output_csv, low_memory=False)
                # Harmonize dtypes on common columns to prevent merge failures
                common = [c for c in df_in.columns if c in df_out.columns]
                for c in common:
                    if df_in[c].dtype != df_out[c].dtype:
                        df_in[c] = df_in[c].astype(str)
                        df_out[c] = df_out[c].astype(str)
                merged = pd.merge(df_in, df_out, on=common, how='outer')
                merged['Scenario'] = scen_name
                # Clean up 'nan' strings back to real NaN
                merged = merged.replace('nan', pd.NA)
                print(f"  Merged: {len(merged)} rows")
                frames.append(merged)
            except Exception as e:
                print(f"  [WARN] Failed: {e}")
        else:
            print(f"[{scen_name}] MISSING — input={input_csv is not None}, output={output_csv is not None}")

    if not frames:
        sys.exit("[ERROR] No scenario data found.")

    combined = pd.concat(frames, ignore_index=True)

    print(f"\n{'='*60}")
    for s in sorted(combined['Scenario'].unique()):
        sub = combined[combined['Scenario'] == s]
        cap = pd.to_numeric(sub.get('TotalCapacityAnnual'), errors='coerce').sum()
        gen = pd.to_numeric(sub.get('ProductionByTechnologyAnnual'), errors='coerce').sum()
        cost = pd.to_numeric(sub.get('TotalDiscountedCost'), errors='coerce').sum()
        print(f"  {s}: {len(sub):,} rows  cap={cap:,.0f}  gen={gen:,.0f}  cost={cost:,.0f}")

    combined.to_csv(args.output, index=False)
    mb = os.path.getsize(args.output) / (1024*1024)
    print(f"\n[DONE] {args.output} ({len(combined):,} rows, {mb:.1f} MB)")

if __name__ == '__main__':
    main()
