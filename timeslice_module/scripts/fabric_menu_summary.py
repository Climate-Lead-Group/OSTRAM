"""
Four-fabric summary table for the training deck.

Reads the four retained workbooks (12 / 16 / 20 / 24 ts) and reports, per
fabric: dominant solar block and its mean CF, total timeslice-indexed row
count, and the solver-size proxy. Solar shape statistics come from
outputs/solar_hour_profile.json if present (built by solar_hour_profile.py).

Usage:  python fabric_menu_summary.py [--json <out.json>]
"""
import argparse
import json
import os
import sys

import pandas as pd

MODULE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(MODULE, 'outputs')

WORKBOOKS = [
    ('3dp/12ts', 'OSTRAM_Timeslice_Outputs_3dp12ts.xlsx', 'this session'),
    ('4dp/16ts', 'OSTRAM_Timeslice_Outputs_4dp16ts.xlsx', 'prior session'),
    ('5dp/20ts', 'OSTRAM_Timeslice_Outputs_REFERENCE_5dp20ts.xlsx', 'ADOPTED'),
    ('6dp/24ts', 'OSTRAM_Timeslice_Outputs_6dp24ts.xlsx', 'this session'),
]
ZONES = ['BGD', 'BTN', 'INDEA', 'INDNE', 'INDNO', 'INDSO', 'INDWE',
         'LKA', 'MDV', 'NPL']


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--json', default=None)
    args = ap.parse_args()

    prof_path = os.path.join(OUT, 'solar_hour_profile.json')
    shape = None
    if os.path.exists(prof_path):
        with open(prof_path, encoding='utf-8') as f:
            shape = json.load(f)

    rows = []
    for name, fn, note in WORKBOOKS:
        path = os.path.join(OUT, fn)
        if not os.path.exists(path):
            print(f"  MISSING: {fn} — reported as NOT COMPUTED")
            rows.append({'fabric': name, 'note': note, 'missing': True})
            continue
        x = pd.ExcelFile(path)
        sheets = {s: x.parse(s) for s in x.sheet_names}
        n_ts = len(sheets['YearSplit'])
        n_dp = n_ts // 4

        indexed = sum(len(sheets[s]) for s in x.sheet_names if s != 'Config')

        frames = []
        for z in ZONES:
            d = sheets[z + '_CF'].copy()
            frames.append(d)
        a = pd.concat(frames, ignore_index=True)
        spv = a[a['tech_code'].astype(str).str.contains('SPV')].copy()
        spv['dp'] = spv['timeslice'].astype(str).str[2:]
        g = spv.groupby('dp')['cf_ninja'].mean()
        dom_dp = g.idxmax()
        dom_lab = str(spv[spv['dp'] == dom_dp]['daypart'].iloc[0])

        rows.append({
            'fabric': name, 'note': note, 'n_dp': n_dp, 'n_ts': n_ts,
            'indexed_rows': indexed,
            'dom_dp': dom_dp, 'dom_label': dom_lab,
            'dom_mean_cf': float(g.max()),
            'workbook_bytes': os.path.getsize(path),
            'workbook': fn,
        })

    base = [r for r in rows if r['fabric'] == '5dp/20ts'][0]

    print("=" * 96)
    print("OSTRAM TIMESLICE FABRIC MENU — four fabrics now available")
    print("=" * 96)
    print(f"\n{'fabric':<10} {'ts':>3} {'dominant solar block':<26} "
          f"{'mean CF':>8} {'vs adopted':>11} {'indexed rows':>13} {'vs adopted':>11}")
    print("-" * 96)
    for r in rows:
        if r.get('missing'):
            print(f"{r['fabric']:<10} {'NOT COMPUTED':>3}")
            continue
        dcf = 100 * (r['dom_mean_cf'] - base['dom_mean_cf']) / base['dom_mean_cf']
        drow = 100 * (r['indexed_rows'] - base['indexed_rows']) / base['indexed_rows']
        mark = '  <-- ADOPTED' if r['note'] == 'ADOPTED' else ''
        print(f"{r['fabric']:<10} {r['n_ts']:>3} "
              f"{r['dom_dp']+' '+r['dom_label']:<26} "
              f"{r['dom_mean_cf']:8.6f} {dcf:+10.2f}% "
              f"{r['indexed_rows']:>13,} {drow:+10.2f}%{mark}")
    print("-" * 96)

    if shape:
        print(f"\n{'fabric':<10} {'phantom solar (dark-hour CF)':<30} "
              f"{'solar-shape RMSE':>17}   {'workbook':>10}")
        print("-" * 96)
        for r in rows:
            if r.get('missing'):
                continue
            f = shape['fabrics'].get(r['fabric'])
            if not f:
                print(f"{r['fabric']:<10} NOT COMPUTED")
                continue
            print(f"{r['fabric']:<10} "
                  f"{f['phantom_dark_cf_hours']:.4f} CF-h/day  "
                  f"= {f['phantom_pct_of_daily']:5.2f}% of daily   "
                  f"{f['rmse']:15.6f}   {r['workbook_bytes']:>9,} B")
        print("-" * 96)
        print("\n  phantom solar = sum(block mean CF x dark hours in block); the")
        print("  capacity a flat-CF block credits to hours the sun is down.")
        print(f"  true daily total = "
              f"{shape['fabrics']['5dp/20ts']['true_daily_cf_hours']:.4f} CF-hours.")
        print("  solar-shape RMSE is UNWEIGHTED and is NOT the sweep's cw_rmse;")
        print("  it does not rank fabrics overall. See VARIANT_12TS_REPORT.md §5.")

    print("\n  Solver-size proxy: timeslice-indexed rows scale exactly with")
    print("  n_timeslices; OSeMOSYS variable count scales with it too, which is")
    print("  the axis on which 6dp/24ts was rejected in favour of 5dp/20ts")
    print("  despite scoring higher (docs/ranking_by_budget.csv: 0.8588 vs 0.8007).")

    if args.json:
        with open(args.json, 'w', encoding='utf-8') as f:
            json.dump({'fabrics': rows}, f, indent=2)
        print(f"\nJSON written: {args.json}")
    return 0


if __name__ == '__main__':
    sys.exit(main())
