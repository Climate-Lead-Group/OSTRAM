"""
rank_timeslice_schemes_v2.py
============================
OSTRAM timeslice scheme selection — ranking v2.

Rationale for v2 changes
------------------------
v2 restructures the composite around three independent information axes
rather than five correlated metrics. The motivation is metric
independence, not a claim about which error type matters more in model
outputs.

At the scheme level (capacity-weighted aggregates across primary
regions), v1 metric correlations were:

       |        | wbv  | rmse | pp   | scv  | wcv  |
       | rmse   | 0.93 | 1.00 | -0.88| -0.10| -0.83|
       | pp     |-0.92 |-0.88 | 1.00 | 0.26 | 0.90 |
       | wcv    |-0.90 |-0.83 | 0.90 |-0.03 | 1.00 |

This reveals three independent axes of information, not five:

  DEMAND AXIS   — captured by RMSE, PP, maxRMSE (all r >= 0.88 with
                  each other). They measure "how well does the step
                  function fit the hourly demand curve."
  SOLAR AXIS    — captured by SCV. r = -0.10 with RMSE; completely
                  orthogonal. Measures "do the block cuts align with
                  the solar availability window." NOT monotonic with
                  block count — 8dp_equal has SCV 1.23, 6dp_D_ramp
                  has SCV 1.81 because its cuts land at solar
                  boundaries (5, 8, 17) rather than at 0, 3, 6, 9...
  WIND AXIS     — captured by WCV. Tracks demand axis at r = -0.83
                  because wind has weak diurnal structure in South
                  Asia, so more blocks helps mechanically. Signal is
                  compressed (0.03-0.06 range across schemes on
                  primary regions) and largely redundant with the
                  demand axis, but not fully zero.
  WBV           — dropped. Spearman rank correlation with RMSE = 0.999.
                  WBV and RMSE rank schemes identically.

Under v1, the demand axis received ~70% of the composite weight
(RMSE 30 + PP 25 + maxRMSE 15) while the orthogonal solar axis
received 20%. This triple-counted the demand signal.

v2 allocates weight by information axis, with each axis receiving
weight proportional to its independent content:

  DEMAND (45%)  RMSE 30 + PP 10 + maxRMSE 5
                RMSE leads because it is the cleanest measure; PP
                catches peak-shape errors that RMSE averages out;
                maxRMSE is a small worst-case tiebreaker.
  SOLAR  (45%)  SCV 45
                Carried by a single metric because there IS only one
                independent metric on this axis.
  WIND   (10%)  WCV 10
                Down-weighted because signal is compressed and
                partially redundant with demand, but retained because
                a small non-zero share breaks ties between otherwise-
                equivalent candidates.

NOTE ON FRAMING
---------------
v2 does NOT assume that demand-fit errors and solar-CF errors have
asymmetric downstream consequences. OSTRAM treats demand as
inelastic within a timeslice (ceteris paribus), so both metric
families flow into investment decisions symmetrically. The v2
reweighting is justified by metric independence, not by differential
error cost.

v1 AND v2 AGREE ON N <= 6
-------------------------
Both weight schemes select 6dp_D_ramp at the N <= 6 budget
(0-5 / 5-8 / 8-17 / 17-20 / 20-22 / 22-24). v2 additionally selects
6dp_D_ramp at the N <= 8 budget, where v1 preferred 8dp_equal.
Under v2 weights, 8dp_equal's superior demand fit no longer
compensates for its poor solar alignment (SCV 1.23 vs 1.81 for
6dp_D_ramp) — demonstrating that more timeslices do not
automatically produce a better scheme when solar alignment is
weighted commensurate with its independent information content.

Open in Spyder and press F5.
Author: CLG / Luis Victor-Gallardo
"""
import os, json
import numpy as np
import pandas as pd

# ======================================================================
# USER CONFIGURATION
# ======================================================================
BASE_DIR    = r"C:\Users\luisfernando\Desktop\OSeMOSYS\asia_ostram\asia_ostram_data"
SENS_DIR    = os.path.join(BASE_DIR, "_sensitivity")
OUTPUT_DIR  = os.path.join(SENS_DIR, "ranking_v2")

SUMMARY_CSV = os.path.join(SENS_DIR, "sensitivity_timeslice_summary.csv")
CONFIG_JSON = os.path.join(SENS_DIR, "sensitivity_config.json")

PRIMARY_REGIONS   = ['BGD', 'LKA', 'INDNO', 'INDEA', 'INDNE', 'INDSO', 'INDWE']
TERTIARY_REGIONS  = ['NPL*', 'BTN*']

CAPACITIES_GW = {
    'BGD':   25,
    'LKA':    5,
    'INDNO': 110,
    'INDEA':  45,
    'INDNE':   6,
    'INDSO':  95,
    'INDWE': 130,
    'NPL*':    3,
    'BTN*':    2,
}

RANKING_SEASON = 'Annual'

# v2 WEIGHTS — three-axis structure
# Each axis receives weight proportional to its independent information content.
#
#   Demand axis (45%) = RMSE 30% + PP 10% + maxRMSE 5%
#   Solar  axis (45%) = SCV  45%
#   Wind   axis (10%) = WCV  10%
WEIGHTS = {
    'scv_score':     0.45,   # SOLAR axis — orthogonal information (r ~ 0 with RMSE)
    'rmse_score':    0.30,   # DEMAND axis — primary demand-fit measure
    'pp_score':      0.10,   # DEMAND axis — peak-shape check on RMSE
    'wcv_score':     0.10,   # WIND axis — compressed but non-zero signal
    'maxrmse_score': 0.05,   # DEMAND axis — worst-region tiebreaker
}
assert abs(sum(WEIGHTS.values()) - 1.0) < 1e-9, "WEIGHTS must sum to 1.0"

# SCV cap — block-width artifacts can inflate CV without adding real
# information. Cap at this value before normalising. Under v2 weights,
# SCV carries 45% so this cap is load-bearing. SCV values of 2.2+
# seen in the B/C candidate families reflect very narrow evening
# blocks rather than genuinely better solar separation. Cap = 2.0
# preserves meaningful discrimination without rewarding geometry.
SCV_CAP = 2.0
WCV_CAP = None

PENALTY_1H_BLOCK = 0.04
PENALTY_2H_BLOCK = 0.01

BUDGETS = [4, 5, 6, 8]
EXCLUDE_SCHEMES = ['12dp_equal', '24dp_hourly']
AGGREGATION = 'capacity'


# ======================================================================
# LOADERS
# ======================================================================
def load_data():
    df = pd.read_csv(SUMMARY_CSV)
    with open(CONFIG_JSON) as f:
        cfg = json.load(f)
    required_cols = {'label', 'season', 'scheme', 'scheme_name',
                     'n_dp', 'wbv', 'pp', 'rmse', 'scv', 'wcv'}
    missing = required_cols - set(df.columns)
    if missing:
        raise ValueError(f"Summary CSV missing columns: {missing}")
    if 'schemes' not in cfg:
        raise ValueError("Config JSON missing 'schemes' key")
    return df, cfg


def block_duration(start, end):
    return (end - start) if end > start else (24 - start + end)


# ======================================================================
# SCORING
# ======================================================================
def compute_aggregates(df, cfg):
    sub = df[(df['season'] == RANKING_SEASON) &
             (df['label'].isin(PRIMARY_REGIONS))].copy()

    n_expected = len(PRIMARY_REGIONS)
    counts = sub.groupby('scheme')['label'].nunique()
    incomplete = counts[counts < n_expected]
    if len(incomplete):
        print(f"[warn] {len(incomplete)} schemes missing some primary regions: "
              f"{list(incomplete.index)[:5]}...")

    def aggregate(group, metric, mode):
        if mode == 'capacity':
            w = group['label'].map(CAPACITIES_GW)
            return float((group[metric] * w).sum() / w.sum())
        else:
            return float(group[metric].mean())

    rows = []
    for scheme, g in sub.groupby('scheme'):
        if scheme in EXCLUDE_SCHEMES:
            continue
        n_dp   = int(g['n_dp'].iloc[0])
        name   = g['scheme_name'].iloc[0]
        blocks = cfg['schemes'].get(scheme, {}).get('blocks', [])

        rows.append({
            'scheme':       scheme,
            'scheme_name':  name,
            'n_dp':         n_dp,
            'n_blocks':     len(blocks),
            'mean_rmse':    aggregate(g, 'rmse', AGGREGATION),
            'mean_pp':      aggregate(g, 'pp',   AGGREGATION),
            'mean_scv':     aggregate(g, 'scv',  AGGREGATION),
            'mean_wcv':     aggregate(g, 'wcv',  AGGREGATION),
            'max_rmse':     float(g['rmse'].max()),
            'worst_region': g.loc[g['rmse'].idxmax(), 'label'],
            'blocks':       blocks,
        })

    return pd.DataFrame(rows)


def normalize_scores(agg):
    a = agg.copy()

    def lower_better(s):
        r = s.max() - s.min()
        return 1 - (s - s.min()) / r if r > 0 else pd.Series(0.5, index=s.index)

    def higher_better(s):
        r = s.max() - s.min()
        return (s - s.min()) / r if r > 0 else pd.Series(0.5, index=s.index)

    scv = a['mean_scv'].clip(upper=SCV_CAP) if SCV_CAP is not None else a['mean_scv']
    wcv = a['mean_wcv'].clip(upper=WCV_CAP) if WCV_CAP is not None else a['mean_wcv']

    a['rmse_score']    = lower_better(a['mean_rmse'])
    a['pp_score']      = higher_better(a['mean_pp'])
    a['scv_score']     = higher_better(scv)
    a['wcv_score']     = higher_better(wcv)
    a['maxrmse_score'] = lower_better(a['max_rmse'])

    return a


def defensibility_penalty(blocks):
    pen = 0.0
    short_blocks = []
    for bname, start, end in blocks:
        dur = block_duration(start, end)
        if dur == 1:
            pen += PENALTY_1H_BLOCK
            short_blocks.append(f"{bname}={dur}h({start}-{end})")
        elif dur == 2:
            pen += PENALTY_2H_BLOCK
            short_blocks.append(f"{bname}={dur}h({start}-{end})")
    return pen, short_blocks


def rank(agg_scored):
    a = agg_scored.copy()
    a['composite'] = sum(a[k] * w for k, w in WEIGHTS.items())

    penalties = a['blocks'].apply(defensibility_penalty)
    a['def_penalty'] = [p for p, _ in penalties]
    a['short_blocks'] = [",".join(sb) if sb else "" for _, sb in penalties]

    a['final'] = a['composite'] - a['def_penalty']
    return a.sort_values('final', ascending=False).reset_index(drop=True)


# ======================================================================
# REPORTING
# ======================================================================
def format_blocks(blocks):
    return " / ".join(f"{s}-{e}" for _, s, e in blocks)


def winners_by_budget(ranked):
    rows = []
    for cap in BUDGETS:
        elig = ranked[ranked['n_dp'] <= cap]
        if len(elig) == 0:
            continue
        top = elig.iloc[0]
        rows.append({
            'budget':       cap,
            'winner':       top['scheme'],
            'n_dp':         int(top['n_dp']),
            'total_ts':     int(top['n_dp'] * 4),
            'final_score':  round(float(top['final']), 4),
            'cw_rmse':      round(float(top['mean_rmse']), 4),
            'cw_pp':        round(float(top['mean_pp']), 3),
            'cw_scv':       round(float(top['mean_scv']), 3),
            'blocks':       format_blocks(top['blocks']),
        })
    return pd.DataFrame(rows)


def write_report(ranked, winners, path):
    lines = []
    lines.append("OSTRAM TIMESLICE SCHEME RANKING — v2 (three-axis)")
    lines.append("=" * 70)
    lines.append(f"Ranking season: {RANKING_SEASON}")
    lines.append(f"Aggregation:    {AGGREGATION} (primary regions only)")
    lines.append(f"Primary:        {', '.join(PRIMARY_REGIONS)}")
    lines.append(f"Excluded:       {', '.join(EXCLUDE_SCHEMES) or '(none)'}")
    lines.append("")
    lines.append("Composite weights (v2 — three-axis structure):")
    lines.append("  Demand axis  (45%): rmse 30 + pp 10 + maxrmse 5")
    lines.append("  Solar  axis  (45%): scv  45")
    lines.append("  Wind   axis  (10%): wcv  10")
    lines.append("  (WBV dropped — Spearman rank corr with RMSE = 0.999)")
    lines.append("")
    lines.append("Individual weights:")
    for k, v in WEIGHTS.items():
        lines.append(f"  {k:18s} {v:.2f}")
    lines.append(f"SCV cap: {SCV_CAP},  WCV cap: {WCV_CAP}")
    lines.append(f"Penalties: 1h block = {PENALTY_1H_BLOCK}, 2h block = {PENALTY_2H_BLOCK}")
    lines.append("")
    lines.append("Rationale: three metric clusters (demand / solar / wind) are")
    lines.append("independent; metrics WITHIN the demand cluster correlate at")
    lines.append("r ~ 0.88 with each other. Weighting by cluster rather than by")
    lines.append("metric count avoids triple-counting the demand signal.")
    lines.append("")

    lines.append("WINNERS BY BUDGET")
    lines.append("-" * 70)
    lines.append(winners.to_string(index=False))
    lines.append("")

    lines.append("FULL RANKING")
    lines.append("-" * 70)
    cols = ['scheme', 'n_dp', 'mean_rmse', 'mean_pp', 'mean_scv', 'max_rmse',
            'worst_region', 'composite', 'def_penalty', 'final', 'short_blocks']
    show = ranked[cols].copy()
    for c in ['mean_rmse', 'mean_pp', 'mean_scv', 'max_rmse',
              'composite', 'def_penalty', 'final']:
        show[c] = show[c].round(4)
    lines.append(show.to_string(index=False))

    lines.append("")
    lines.append("BLOCK DEFINITIONS (top 10 by final score)")
    lines.append("-" * 70)
    for _, r in ranked.head(10).iterrows():
        lines.append(f"  {r['scheme']:22s}  {format_blocks(r['blocks'])}")

    text = "\n".join(lines)
    with open(path, 'w', encoding='utf-8') as f:
        f.write(text)
    return text


# ======================================================================
# MAIN
# ======================================================================
def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    print(f"Loading {SUMMARY_CSV}")
    df, cfg = load_data()
    print(f"  {len(df)} rows, {df['scheme'].nunique()} schemes, "
          f"{df['label'].nunique()} regions, {df['season'].nunique()} seasons")

    print("\nComputing aggregates...")
    agg = compute_aggregates(df, cfg)
    print(f"  {len(agg)} candidate schemes after exclusions")

    print("Normalizing + ranking (v2 three-axis weights)...")
    scored = normalize_scores(agg)
    ranked = rank(scored)

    winners = winners_by_budget(ranked)

    ranked_out = ranked.drop(columns=['blocks']).copy()
    ranked_out.to_csv(os.path.join(OUTPUT_DIR, 'ranking_full.csv'), index=False)
    winners.to_csv(os.path.join(OUTPUT_DIR, 'ranking_by_budget.csv'), index=False)

    report_text = write_report(ranked, winners, os.path.join(OUTPUT_DIR, 'ranking_report.txt'))

    print("\n" + "=" * 70)
    print(report_text)
    print("=" * 70)
    print(f"\nWritten to {OUTPUT_DIR}")


main()
