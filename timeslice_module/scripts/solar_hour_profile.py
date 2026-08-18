"""
Build the hour-resolved solar CF truth from the Renewables.ninja hourly table,
then score every candidate daypart fabric against it.

Why this exists. The variant brief expected the 3dp/12ts fabric to show a
*diluted* solar block mean versus the adopted 5dp/20ts. It does the opposite
(+24% sharper), because 08-16 is a NARROWER window around solar noon than
06-17 and so excludes the weak dawn/dusk shoulders. Block mean CF is therefore
the wrong statistic for "does this fabric lose the solar peak" — it rewards
narrow blocks regardless of whether the rest of the day is represented.

The right statistic is how well the fabric's piecewise-constant CF reproduces
the true hourly shape. That is computed here as RMSE, plus the quantity that
actually misleads a solver: solar capacity credited to hours that are dark.

The hourly truth does not depend on DAYPART_DEF, so it is computed once and
all four fabrics are scored against the same baseline.

Usage:
    python solar_hour_profile.py [--hourly <compiled_reninja_hourly.csv>]
                                 [--json <out.json>]
"""
import argparse
import json
import os
import sys

import numpy as np
import pandas as pd

MODULE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_HOURLY = os.path.join(
    MODULE, 'inputs', '_Reno_Ninja', 'ninja_data', 'output_rebuilt',
    'compiled_reninja_hourly.csv')

USECOLS = ['cf', 'resource', 'region', 'hour_local', 'season', 'model_year']

# The rebuilder drops model_years with < 8000 hours as incomplete; mirror that
# so this profile is comparable with compiled_reninja_ts.csv.
MIN_HOURS = 8000

FABRICS = {
    '3dp/12ts': [('D1', 'Night', 0, 8), ('D2', 'Day', 8, 16),
                 ('D3', 'Evening', 16, 24)],
    '4dp/16ts': [('D1', 'Night', 0, 6), ('D2', 'Morning', 6, 12),
                 ('D3', 'Afternoon', 12, 18), ('D4', 'Evening', 18, 24)],
    '5dp/20ts': [('D1', 'Night', 0, 6), ('D2', 'Solar day', 6, 17),
                 ('D3', 'Evening peak', 17, 20), ('D4', 'Late evening', 20, 22),
                 ('D5', 'Late night', 22, 24)],
    '6dp/24ts': [('D1', 'Night', 0, 5), ('D2', 'Dawn', 5, 8),
                 ('D3', 'Solar day', 8, 17), ('D4', 'Evening peak', 17, 20),
                 ('D5', 'Late evening', 20, 22), ('D6', 'Late night', 22, 24)],
}


def load_hour_profile(path):
    """mean solar cf per (region, season, hour_local), matching the ts pipeline."""
    print(f"Reading {os.path.basename(path)} "
          f"({os.path.getsize(path)/1e6:.0f} MB, 6 of 36 columns)...")
    # Pass 1: hours per (region, model_year) to apply the completeness filter.
    counts = {}
    sums = {}
    n = 0
    for chunk in pd.read_csv(path, usecols=USECOLS, chunksize=2_000_000):
        chunk = chunk[chunk['resource'] == 'solar']
        n += len(chunk)
        for key, g in chunk.groupby(['region', 'model_year'], sort=False):
            counts[key] = counts.get(key, 0) + len(g)
        for key, g in chunk.groupby(['region', 'season', 'hour_local',
                                     'model_year'], sort=False):
            s, c = sums.get(key, (0.0, 0))
            sums[key] = (s + g['cf'].sum(), c + len(g))
        print(f"  {n:,} solar rows so far")

    complete = {k for k, v in counts.items() if v >= MIN_HOURS}
    dropped = sorted(k for k in counts if k not in complete)
    print(f"  model_year completeness filter (>= {MIN_HOURS} h): "
          f"{len(complete)} kept, {len(dropped)} dropped")
    if dropped:
        yrs = sorted({k[1] for k in dropped})
        print(f"    dropped model_years: {yrs}")

    rec = []
    for (reg, sea, hr, my), (s, c) in sums.items():
        if (reg, my) in complete and c:
            rec.append({'region': reg, 'season': sea, 'hour': int(hr),
                        'model_year': my, 'cf': s / c})
    df = pd.DataFrame(rec)
    # mean over model_years -> (region, season, hour), mirroring the generator's
    # groupby(resource, region, timeslice).mean() over model_year.
    prof = (df.groupby(['region', 'season', 'hour'])['cf']
              .mean().reset_index())
    print(f"  profile: {len(prof)} (region, season, hour) cells, "
          f"{prof['region'].nunique()} regions, {prof['season'].nunique()} seasons")
    return prof


def block_mean(prof, start, end):
    """Fabric block mean, aggregated the same way the check suite does:
    per (region, season) block mean first, then unweighted mean over the 40
    (region, season) pairs — matching cf_ninja averaged over 10 zones x 4 seasons."""
    sub = prof[(prof['hour'] >= start) & (prof['hour'] < end)]
    per_cell = sub.groupby(['region', 'season'])['cf'].mean()
    return float(per_cell.mean()), per_cell


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--hourly', default=DEFAULT_HOURLY)
    ap.add_argument('--json', default=None)
    args = ap.parse_args()

    prof = load_hour_profile(args.hourly)
    out = {'source': os.path.basename(args.hourly)}

    # ---- the hourly truth, pooled over regions and seasons ----
    pooled = prof.groupby('hour')['cf'].mean()
    print("\n" + "=" * 74)
    print("TRUE MEAN SOLAR CF BY LOCAL HOUR  (10 zones x 4 seasons pooled)")
    print("=" * 74)
    for h in range(24):
        v = pooled.get(h, 0.0)
        bar = '#' * int(round(v * 60))
        print(f"  {h:02d}-{h+1:02d}  {v:8.6f}  {bar}")
    daylight = [h for h in range(24) if pooled.get(h, 0) > 0.005]
    print(f"\n  hours with CF > 0.005 : {daylight[0]:02d}-{daylight[-1]+1:02d}"
          f"  ({len(daylight)} h)")
    print(f"  peak hour             : {int(pooled.idxmax()):02d}-"
          f"{int(pooled.idxmax())+1:02d}  CF {pooled.max():.6f}")
    dark = [h for h in range(24) if pooled.get(h, 0) <= 0.005]
    print(f"  dark hours (CF<=0.005): {len(dark)}  {dark}")
    out['hourly_truth'] = {int(h): float(pooled.get(h, 0.0)) for h in range(24)}
    out['dark_hours'] = dark

    # ---- score each fabric ----
    print("\n" + "=" * 74)
    print("FABRIC SCORING vs THE HOURLY TRUTH")
    print("=" * 74)
    fab_out = {}
    for name, dp in FABRICS.items():
        # piecewise-constant reconstruction, per (region, season)
        cells = prof.groupby(['region', 'season'])
        sse, n_pts = 0.0, 0
        for (reg, sea), g in cells:
            truth = g.set_index('hour')['cf'].reindex(range(24)).fillna(0.0)
            approx = pd.Series(0.0, index=range(24))
            for _c, _l, s, e in dp:
                approx.loc[s:e - 1] = truth.loc[s:e - 1].mean()
            sse += float(((truth - approx) ** 2).sum())
            n_pts += 24
        rmse = (sse / n_pts) ** 0.5

        blocks = []
        for c, l, s, e in dp:
            bm, _ = block_mean(prof, s, e)
            n_dark = sum(1 for h in range(s, e) if h in dark)
            blocks.append({'code': c, 'label': l, 'start': s, 'end': e,
                           'hours': e - s, 'mean_cf': bm,
                           'dark_hours_in_block': n_dark})
        dom = max(blocks, key=lambda b: b['mean_cf'])

        # Solar credited to genuinely dark hours: the fabric asserts a flat CF
        # across its block, so every dark hour inside a block with mean>0 is
        # handed capacity that physically is not there.
        phantom = sum(b['mean_cf'] * b['dark_hours_in_block'] for b in blocks)
        true_daily = float(pooled.reindex(range(24)).fillna(0.0).sum())

        fab_out[name] = {'blocks': blocks, 'rmse': rmse,
                         'dominant': dom,
                         'phantom_dark_cf_hours': phantom,
                         'true_daily_cf_hours': true_daily,
                         'phantom_pct_of_daily': 100 * phantom / true_daily}
        print(f"\n  {name}")
        print(f"    {'block':<6}{'hours':<9}{'h':>3}  {'mean CF':>9}  dark h")
        for b in blocks:
            print(f"    {b['code']:<6}{b['start']:02d}-{b['end']:02d}    "
                  f"{b['hours']:>3}  {b['mean_cf']:9.6f}  "
                  f"{b['dark_hours_in_block']:>3}")
        print(f"    dominant block      : {dom['code']} "
              f"({dom['start']:02d}-{dom['end']:02d}) mean {dom['mean_cf']:.6f}")
        print(f"    RMSE vs hourly truth: {rmse:.6f}")
        print(f"    phantom solar       : {phantom:.4f} CF-hours/day credited "
              f"to dark hours = {100*phantom/true_daily:.2f}% of the true "
              f"daily {true_daily:.4f} CF-hours")

    # ---- comparison table ----
    print("\n" + "=" * 74)
    print("SUMMARY  (dominant solar block; RMSE lower = better shape fidelity)")
    print("=" * 74)
    print(f"  {'fabric':<10} {'dom':<5} {'window':<8} {'mean CF':>9} "
          f"{'vs 20ts':>9} {'RMSE':>9} {'phantom %':>10}")
    base = fab_out['5dp/20ts']['dominant']['mean_cf']
    for name in ['3dp/12ts', '4dp/16ts', '5dp/20ts', '6dp/24ts']:
        f = fab_out[name]
        d = f['dominant']
        rel = 100 * (d['mean_cf'] - base) / base
        print(f"  {name:<10} {d['code']:<5} "
              f"{d['start']:02d}-{d['end']:02d}    {d['mean_cf']:9.6f} "
              f"{rel:+8.2f}% {f['rmse']:9.6f} "
              f"{f['phantom_pct_of_daily']:9.2f}%")
    out['fabrics'] = fab_out

    if args.json:
        with open(args.json, 'w', encoding='utf-8') as fh:
            json.dump(out, fh, indent=2)
        print(f"\nJSON written: {args.json}")
    return 0


if __name__ == '__main__':
    sys.exit(main())
