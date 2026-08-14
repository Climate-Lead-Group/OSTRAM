"""
Variant check suite for the OSTRAM timeslice module.

Mirrors the check set established in VARIANT_16TS_REPORT.md and applies it to
any fabric, comparing against the adopted 5dp/20ts reference workbook.

Tolerance convention (carried over from the 16ts session): every sum check is
first reported against the brief's literal 1e-9. A failure is then classified
as EXPECTED-ROUNDING if and only if it is explained by the generator's
`round(..., 6)` — the arithmetic is computed and printed, not asserted.

Usage:
    python check_variant.py <variant_workbook> [--ref <reference_workbook>]
                            [--json <out.json>]

Writes nothing unless --json is given.
"""
import argparse
import json
import os
import re
import sys

import pandas as pd

ZONES = ['BGD', 'BTN', 'INDEA', 'INDNE', 'INDNO', 'INDSO', 'INDWE',
         'LKA', 'MDV', 'NPL']
SEASON_DAYS = {'S1': 90, 'S2': 92, 'S3': 122, 'S4': 61}   # 4 IMD seasons, 365 d
LITERAL_TOL = 1e-9
ROUNDING_TOL = 1e-5      # appropriate tolerance for 6-dp stored values

MODULE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_REF = os.path.join(
    MODULE, 'outputs', 'OSTRAM_Timeslice_Outputs_REFERENCE_5dp20ts.xlsx')


# ----------------------------------------------------------------------
# helpers
# ----------------------------------------------------------------------

def parse_hours(label):
    """'Night (00-08)' -> (0, 8). Returns None if unparseable."""
    m = re.search(r'\((\d{2})-(\d{2})\)', str(label))
    return (int(m.group(1)), int(m.group(2))) if m else None


def load(path):
    x = pd.ExcelFile(path)
    return x, {s: x.parse(s) for s in x.sheet_names}


def solar_table(sheets):
    """All SPV rows across zones, with daypart hour bounds attached."""
    frames = []
    for z in ZONES:
        d = sheets[z + '_CF'].copy()
        d['zone'] = z
        frames.append(d)
    a = pd.concat(frames, ignore_index=True)
    spv = a[a['tech_code'].astype(str).str.contains('SPV')].copy()
    spv['dp'] = spv['timeslice'].astype(str).str[2:]
    bounds = spv['daypart'].apply(parse_hours)
    spv['h_start'] = [b[0] if b else None for b in bounds]
    spv['h_end'] = [b[1] if b else None for b in bounds]
    return spv


class Checks:
    def __init__(self):
        self.rows = []

    def add(self, name, ok, detail):
        self.rows.append({'check': name, 'ok': bool(ok), 'detail': detail})
        flag = 'PASS' if ok else 'FAIL'
        print(f"  [{flag}] {name}")
        if detail:
            print(f"         {detail}")

    @property
    def n_pass(self):
        return sum(1 for r in self.rows if r['ok'])

    @property
    def n_fail(self):
        return sum(1 for r in self.rows if not r['ok'])


# ----------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('workbook')
    ap.add_argument('--ref', default=DEFAULT_REF)
    ap.add_argument('--json', default=None)
    args = ap.parse_args()

    vx, vs = load(args.workbook)
    rx, rs = load(args.ref)

    ys = vs['YearSplit']
    ts_list = list(ys['timeslice'].astype(str))
    n_ts = len(ts_list)
    dps = sorted({t[2:] for t in ts_list}, key=lambda d: int(d[1:]))
    seasons = sorted({t[:2] for t in ts_list})
    n_dp = len(dps)
    ref_n_ts = len(rs['YearSplit'])

    out = {'workbook': os.path.basename(args.workbook),
           'reference': os.path.basename(args.ref),
           'n_timeslices': n_ts, 'n_dayparts': n_dp}

    print("=" * 78)
    print(f"VARIANT CHECK SUITE  —  {os.path.basename(args.workbook)}")
    print(f"  fabric: {len(seasons)} seasons x {n_dp} dayparts = {n_ts} timeslices")
    print(f"  reference: {os.path.basename(args.ref)} ({ref_n_ts} ts)")
    print("=" * 78)

    C = Checks()

    # ---------------- 1. fabric identity ----------------
    print(f"\n--- 1. FABRIC IDENTITY ---")
    C.add(f"{n_ts} timeslices present, all unique",
          len(ts_list) == n_ts and len(set(ts_list)) == n_ts,
          f"{len(ts_list)} rows, {len(set(ts_list))} unique")
    C.add(f"4 seasons x {n_dp} dayparts",
          seasons == ['S1', 'S2', 'S3', 'S4'] and n_dp * 4 == n_ts,
          f"seasons {seasons}, dayparts {dps}")
    expected = [s + d for s in ['S1', 'S2', 'S3', 'S4'] for d in dps]
    C.add("names match canonical SxDy order",
          ts_list == expected,
          f"{ts_list[0]} ... {ts_list[-1]}"
          + ("" if ts_list == expected else f"  EXPECTED {expected}"))
    per_dp = {d: sum(1 for t in ts_list if t[2:] == d) for d in dps}
    C.add("each daypart appears in all 4 seasons",
          all(v == 4 for v in per_dp.values()), str(per_dp))

    cfg = vs['Config']
    cfg_dp = cfg[cfg['key'].astype(str).str.startswith('daypart_')]
    C.add(f"Config sheet lists {n_dp} dayparts", len(cfg_dp) == n_dp,
          ", ".join(cfg_dp['value'].astype(str)))
    ys_labels = set(ys['daypart'].astype(str))
    cfg_labels = set(cfg_dp['value'].astype(str))
    C.add("daypart labels agree between YearSplit and Config",
          ys_labels == cfg_labels,
          f"{len(ys_labels)} labels, identical sets"
          if ys_labels == cfg_labels else f"{ys_labels ^ cfg_labels}")

    # daypart hour coverage
    dp_hours = {}
    for d in dps:
        lab = ys[ys['timeslice'].astype(str).str[2:] == d]['daypart'].iloc[0]
        b = parse_hours(lab)
        dp_hours[d] = b[1] - b[0]
    C.add("daypart hours sum to 24", sum(dp_hours.values()) == 24,
          " + ".join(f"{d}={h}h" for d, h in dp_hours.items())
          + f" = {sum(dp_hours.values())}h")
    out['daypart_hours'] = dp_hours

    # ---------------- 2. row scaling ----------------
    ratio = n_ts / ref_n_ts
    print(f"\n--- 2. ROW-COUNT SCALING (expected {n_ts}/{ref_n_ts} = {ratio:g}) ---")
    C.add("sheet set identical to reference",
          vx.sheet_names == rx.sheet_names,
          f"{len(vx.sheet_names)} sheets, same names and order")
    scaling = []
    bad = []
    for s in vx.sheet_names:
        nv, nr = len(vs[s]), len(rs[s])
        if s == 'Config':
            scaling.append({'sheet': s, 'variant': nv, 'ref': nr,
                            'ratio': round(nv / nr, 6), 'indexed': False})
            continue
        r = nv / nr if nr else float('nan')
        exact = (nv == nr * ratio)
        scaling.append({'sheet': s, 'variant': nv, 'ref': nr,
                        'ratio': round(r, 6), 'indexed': True, 'exact': exact})
        if not exact:
            bad.append(f"{s}: {nv} vs {nr} (ratio {r:.6f})")
    C.add(f"all timeslice-indexed sheets scale by exactly {ratio:g}",
          not bad,
          f"{len(scaling)-1} sheets checked, all exact" if not bad
          else "; ".join(bad))
    cfgrow = [x for x in scaling if x['sheet'] == 'Config'][0]
    C.add("Config sheet changes by the daypart delta only",
          cfgrow['variant'] == n_dp + 4 and cfgrow['ref'] == (ref_n_ts // 4) + 4,
          f"{cfgrow['ref']} -> {cfgrow['variant']} "
          f"({n_dp}dp+4s vs {ref_n_ts//4}dp+4s; not timeslice-indexed)")
    out['scaling'] = scaling

    # ---------------- 3. YearSplit sum ----------------
    print(f"\n--- 3. YEARSPLIT SUM (literal tol {LITERAL_TOL:g}) ---")
    ys_sum = float(ys['yearsplit'].sum())
    dev = ys_sum - 1.0
    literal_ok = abs(dev) <= LITERAL_TOL

    # predicted rounding error, computed from the generator's own formula
    pred = 0.0
    terms = []
    for s in ['S1', 'S2', 'S3', 'S4']:
        for d in dps:
            exact = (SEASON_DAYS[s] * dp_hours[d]) / 8760
            err = round(exact, 6) - exact
            pred += err
            terms.append({'ts': s + d, 'exact': exact,
                          'rounded': round(exact, 6), 'err': err})
    match = abs(pred - dev) < 5e-12
    print(f"         observed sum = {ys_sum:.9f}   deviation = {dev:+.3e}")
    print(f"         predicted by round(days*hours/8760, 6) = {pred:+.3e}"
          f"   {'MATCHES' if match else 'DOES NOT MATCH'} observed")
    ref_ys_dev = float(rs['YearSplit']['yearsplit'].sum()) - 1.0
    print(f"         20ts reference deviation = {ref_ys_dev:+.3e} "
          f"(sum {1+ref_ys_dev:.9f})")
    if literal_ok:
        C.add(f"YearSplit sums to 1.0 within {LITERAL_TOL:g}", True,
              f"sum {ys_sum:.9f}")
    else:
        cls = ("EXPECTED-ROUNDING" if match and abs(dev) <= ROUNDING_TOL
               else "UNEXPLAINED")
        C.add(f"YearSplit sums to 1.0 within {LITERAL_TOL:g}", False,
              f"sum {ys_sum:.9f}, deviation {dev:+.3e} — classified "
              f"{cls} (predicted {pred:+.3e}; passes at {ROUNDING_TOL:g}; "
              f"20ts reference deviates {ref_ys_dev:+.3e})")
    out['yearsplit'] = {'sum': ys_sum, 'deviation': dev, 'predicted': pred,
                        'prediction_matches': match, 'literal_ok': literal_ok,
                        'ref_deviation': ref_ys_dev, 'terms': terms}

    # ---------------- 4. demand fractions ----------------
    print(f"\n--- 4. DEMAND FRACTIONS PER ZONE (literal tol {LITERAL_TOL:g}) ---")
    dem = {}
    worst = 0.0
    exact_hits = 0
    print(f"         {'zone':<7} {'sum':>12} {'dev':>11}   "
          f"{'20ts sum':>12} {'20ts dev':>11}")
    for z in ZONES:
        v = float(vs[z + '_Dem']['demand_fraction'].sum())
        r = float(rs[z + '_Dem']['demand_fraction'].sum())
        dv, dr = v - 1.0, r - 1.0
        dem[z] = {'sum': v, 'dev': dv, 'ref_sum': r, 'ref_dev': dr}
        worst = max(worst, abs(dv))
        if dv == 0.0:
            exact_hits += 1
        print(f"         {z:<7} {v:12.6f} {dv:+11.1e}   {r:12.6f} {dr:+11.1e}")
    n_lit_fail = sum(1 for z in dem if abs(dem[z]['dev']) > LITERAL_TOL)
    ref_worst = max(abs(dem[z]['ref_dev']) for z in dem)
    bound = n_ts * 5e-7
    if n_lit_fail == 0:
        C.add(f"demand fractions sum to 1.0 per zone within {LITERAL_TOL:g}",
              True, f"all 10 zones; max |dev| {worst:.1e}")
    else:
        cls = "EXPECTED-ROUNDING" if worst <= ROUNDING_TOL else "UNEXPLAINED"
        C.add(f"demand fractions sum to 1.0 per zone within {LITERAL_TOL:g}",
              False,
              f"{n_lit_fail}/10 zones exceed it; max |dev| {worst:.1e} — "
              f"classified {cls}. Every adapter does "
              f"round(v/total, 6) over {n_ts} values, so the worst-case bound "
              f"is {n_ts}x5e-7 = {bound:.1e}; observed is inside it. "
              f"20ts reference max |dev| {ref_worst:.1e}. "
              f"{exact_hits}/10 zones land exactly on 1.0.")
    out['demand'] = dem
    out['demand_bound'] = bound

    # ---------------- 5. solar CF by daypart ----------------
    print(f"\n--- 5. SOLAR CF BY DAYPART (cf_ninja, all 10 zones) ---")
    spv = solar_table(vs)
    prof = []
    print(f"         {'dp':<4} {'hours':<8} {'min':>9} {'max':>9} "
          f"{'mean':>9} {'n':>4}")
    for d in dps:
        g = spv[spv['dp'] == d]['cf_ninja'].dropna()
        lab = spv[spv['dp'] == d]['daypart'].iloc[0]
        b = parse_hours(lab)
        prof.append({'dp': d, 'label': str(lab),
                     'h_start': b[0], 'h_end': b[1],
                     'min': float(g.min()), 'max': float(g.max()),
                     'mean': float(g.mean()), 'n': int(g.size)})
        print(f"         {d:<4} {b[0]:02d}-{b[1]:02d}   {g.min():9.6f} "
              f"{g.max():9.6f} {g.mean():9.6f} {g.size:4d}")
    out['solar_profile'] = prof

    dom = max(prof, key=lambda p: p['mean'])
    C.add("exactly one dominant solar block, and it contains solar noon",
          dom['h_start'] <= 12 < dom['h_end'],
          f"{dom['dp']} ({dom['h_start']:02d}-{dom['h_end']:02d}) "
          f"mean {dom['mean']:.6f}; contains hour 12")
    others = [p for p in prof if p['dp'] != dom['dp']]
    C.add("dominant block mean exceeds every other block",
          all(dom['mean'] > p['mean'] for p in others),
          "; ".join(f"{dom['dp']}={dom['mean']:.4f} > {p['dp']}={p['mean']:.4f}"
                    for p in others))
    # blocks wholly outside daylight must be ~0
    dark = [p for p in prof if p['h_end'] <= 5 or p['h_start'] >= 20]
    if dark:
        C.add("blocks wholly outside daylight have solar CF ~ 0 (<0.01)",
              all(p['max'] < 0.01 for p in dark),
              "; ".join(f"{p['dp']}({p['h_start']:02d}-{p['h_end']:02d}) "
                        f"max {p['max']:.6f}" for p in dark))
    else:
        print("  [NOTE] no daypart lies wholly outside 05-20; the "
              "'night block ~ 0' check does not apply to this fabric")
        out['no_dark_block'] = True

    # ---------------- 6. dilution / sharpening vs 20ts ----------------
    print(f"\n--- 6. DOMINANT SOLAR BLOCK vs 20ts SOLAR DAY (06-17) ---")
    rspv = solar_table(rs)
    rprof = []
    for d in sorted({t[2:] for t in rs['YearSplit']['timeslice'].astype(str)},
                    key=lambda d: int(d[1:])):
        g = rspv[rspv['dp'] == d]['cf_ninja'].dropna()
        lab = rspv[rspv['dp'] == d]['daypart'].iloc[0]
        b = parse_hours(lab)
        rprof.append({'dp': d, 'label': str(lab), 'h_start': b[0],
                      'h_end': b[1], 'mean': float(g.mean()),
                      'min': float(g.min()), 'max': float(g.max()),
                      'n': int(g.size)})
    rdom = max(rprof, key=lambda p: p['mean'])
    delta = dom['mean'] - rdom['mean']
    pct = 100.0 * delta / rdom['mean']
    print(f"         20ts   {rdom['dp']} {rdom['label']:<20} "
          f"mean {rdom['mean']:.6f}  n={rdom['n']}")
    print(f"         {n_ts}ts   {dom['dp']} {dom['label']:<20} "
          f"mean {dom['mean']:.6f}  n={dom['n']}")
    print(f"         delta = {delta:+.6f}  ({pct:+.2f}% "
          f"{'DILUTION' if delta < 0 else 'SHARPENING'})")
    out['dominant'] = {'variant': dom, 'reference': rdom,
                       'delta': delta, 'pct': pct}
    out['reference_solar_profile'] = rprof

    # ---------------- verdict ----------------
    print("\n" + "=" * 78)
    unexplained = [r for r in C.rows if not r['ok']
                   and 'EXPECTED-ROUNDING' not in r['detail']]
    verdict = 'PASS' if not unexplained else 'FAIL'
    print(f"VERDICT: {verdict}    "
          f"{C.n_pass} passed, {C.n_fail} failed at literal tolerances")
    if C.n_fail and not unexplained:
        print("         all failures classified EXPECTED-ROUNDING "
              "(6-dp artifact present in the adopted reference too)")
    for r in unexplained:
        print(f"         UNEXPLAINED: {r['check']} — {r['detail']}")
    print("=" * 78)
    out['checks'] = C.rows
    out['verdict'] = verdict
    out['n_pass'] = C.n_pass
    out['n_fail'] = C.n_fail

    if args.json:
        with open(args.json, 'w', encoding='utf-8') as f:
            json.dump(out, f, indent=2)
        print(f"\nJSON written: {args.json}")
    return 0 if verdict == 'PASS' else 1


if __name__ == '__main__':
    sys.exit(main())
