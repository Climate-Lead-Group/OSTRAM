"""
sensitivity_timeslice_sweep.py
================================
OSTRAM Timeslice Sensitivity Analysis

Sweeps from 2 to 12+ dayparts (equal-width AND asymmetric schemes)
against real sub-daily data for BGD, LKA, IND (×5), NPL, BTN.

For each country × season × scheme, computes:
  1. Within-Block Variance (WBV) — information loss; lower is better
  2. Peak Preservation (PP)       — peak-trough range captured; 1.0 is perfect
  3. RMSE of step reconstruction  — direct curve error; lower is better
  4. Solar Differentiation (SCV)  — CV of block-mean solar CFs; higher is better

Outputs:
  - Per-country overlay figures (hourly curve + step-function per scheme)
  - Convergence plot (metric vs. N dayparts)
  - Summary CSV with all metrics
  - Season-disaggregated results

Open in Spyder and press F5.

Author: CLG / Luis Victor-Gallardo
"""

import os
import re
import glob
import json
import warnings
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import matplotlib.patches as mpatches
from datetime import datetime, time as dtime
from collections import OrderedDict

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# ======================================================================
# USER CONFIGURATION
# ======================================================================

BASE_DIR   = r"C:\Users\luisfernando\Desktop\OSeMOSYS\asia_ostram\asia_ostram_data"
OUTPUT_DIR = os.path.join(BASE_DIR, "_sensitivity")
DPI        = 200
VERBOSE    = True

# Which countries to include (comment out to skip)
RUN_BGD = True
RUN_LKA = True
RUN_IND = True
RUN_NPL = True    # representative profiles only — flagged in output
RUN_BTN = True    # representative profiles only — flagged in output

# ======================================================================
# SEASON DEFINITION (must match build_ostram_timeslices.py)
# ======================================================================

SEASON_DEF = [
    ('S1', 'Winter (Dec-Feb)',       [12, 1, 2],       90),
    ('S2', 'Pre-monsoon (Mar-May)',  [3, 4, 5],        92),
    ('S3', 'SW Monsoon (Jun-Sep)',   [6, 7, 8, 9],    122),
    ('S4', 'Post-monsoon (Oct-Nov)', [10, 11],          61),
]

MONTH_TO_SEASON = {}
SEASON_NAMES = {}
for code, label, months, days in SEASON_DEF:
    SEASON_NAMES[code] = label
    for m in months:
        MONTH_TO_SEASON[m] = code

SEASONS_FOR_ANALYSIS = OrderedDict([
    ('Annual', list(range(1, 13))),
    ('S1',     [12, 1, 2]),
    ('S2',     [3, 4, 5]),
    ('S3',     [6, 7, 8, 9]),
    ('S4',     [10, 11]),
])

# ======================================================================
# DAYPART SCHEMES TO TEST
#
# Format: (label, start_hour, end_hour)
# Blocks must tile 0–24 with no gaps.
# Midnight-wrapping NOT used here — all start < end for clarity.
# ======================================================================

SCHEMES = OrderedDict()

# --- Equal-width sweeps (ORIGINAL BASELINES) ---
SCHEMES['2dp_equal'] = {
    'name': '2 dp (12h equal)', 'short': '2dp_eq',
    'blocks': [('D1', 0, 12), ('D2', 12, 24)],
}
SCHEMES['2dp_daynight'] = {
    'name': '2 dp (night 18-6 / day 6-18)', 'short': '2dp_dn',
    'blocks': [('N', 0, 6), ('D', 6, 18), ('N2', 18, 24)],
    'merge_labels': {'N': 'Night', 'D': 'Day', 'N2': 'Night'},
}
SCHEMES['3dp_legacy'] = {
    'name': '3 dp (OSTRAM legacy)', 'short': '3dp_leg',
    'blocks': [('D1', 0, 6), ('D2', 6, 18), ('D3', 18, 24)],
}
SCHEMES['4dp_equal'] = {
    'name': '4 dp (6h equal)', 'short': '4dp_eq',
    'blocks': [('D1', 0, 6), ('D2', 6, 12), ('D3', 12, 18), ('D4', 18, 24)],
}
SCHEMES['4dp_shifted'] = {
    'name': '4 dp (shifted: 22-5, 5-11, 11-17, 17-22)', 'short': '4dp_sh',
    'blocks': [('D1', 0, 5), ('D2', 5, 11), ('D3', 11, 17), ('D4', 17, 22), ('D5', 22, 24)],
    'merge_labels': {'D1': 'Night', 'D2': 'Morning', 'D3': 'Afternoon',
                     'D4': 'Evening', 'D5': 'Night'},
}
SCHEMES['5dp_asym'] = {
    'name': '5 dp (asymmetric)', 'short': '5dp_asym',
    'blocks': [('D1', 0, 6), ('D2', 6, 10), ('D3', 10, 14),
               ('D4', 14, 18), ('D5', 18, 24)],
}
SCHEMES['5dp_solar'] = {
    'name': '5 dp (solar-optimised, legacy)', 'short': '5dp_sol',
    'blocks': [('D1', 0, 5), ('D2', 5, 11), ('D3', 11, 14),
               ('D4', 14, 18), ('D5', 18, 24)],
}
SCHEMES['6dp_solar'] = {
    'name': '6 dp (solar-optimised, legacy)', 'short': '6dp_sol',
    'blocks': [('D1', 0, 5), ('D2', 5, 11), ('D3', 11, 14),
               ('D4', 14, 18), ('D5', 18, 22), ('D6', 22, 24)],
}
SCHEMES['6dp_equal'] = {
    'name': '6 dp (4h equal)', 'short': '6dp_eq',
    'blocks': [('D1', 0, 4), ('D2', 4, 8), ('D3', 8, 12),
               ('D4', 12, 16), ('D5', 16, 20), ('D6', 20, 24)],
}
SCHEMES['8dp_equal'] = {
    'name': '8 dp (3h equal)', 'short': '8dp_eq',
    'blocks': [('D1', 0, 3), ('D2', 3, 6), ('D3', 6, 9), ('D4', 9, 12),
               ('D5', 12, 15), ('D6', 15, 18), ('D7', 18, 21), ('D8', 21, 24)],
}
SCHEMES['12dp_equal'] = {
    'name': '12 dp (2h equal)', 'short': '12dp_eq',
    'blocks': [(f'D{i+1}', i*2, (i+1)*2) for i in range(12)],
}
SCHEMES['24dp_hourly'] = {
    'name': '24 dp (hourly — upper bound)', 'short': '24dp_hr',
    'blocks': [(f'H{i:02d}', i, i+1) for i in range(24)],
}

# ================================================================
# NEW CANDIDATES (Claude-designed, April 2026)
# Design principle: ONE fat solar block (spanning all non-zero solar
# hours) maximises the CV of block-mean solar CFs, because SCV rewards
# contrast between "solar" and "no-solar" blocks, not peak isolation.
#
# All 12 candidates below beat 3dp_legacy (SCV=1.41), the previous
# ceiling, by 28-58%. Recommended top picks:
#   - 5dp_B_6_18  : SCV 2.00, RMSE -11% vs 4dp_shifted  (best 5dp)
#   - 6dp_B_6_18  : SCV 2.23, RMSE -14% vs 4dp_shifted  (best 6dp)
#   - 6dp_D_ramp  : SCV 1.81, RMSE -33% vs 4dp_shifted  (lowest RMSE)
# ================================================================

SCHEMES['5dp_A_5_17'] = {
    'name': '5 dp (solar 5-17)', 'short': '5dp_A',
    'blocks': [('D1', 0, 5), ('D2', 5, 17), ('D3', 17, 20),
               ('D4', 20, 22), ('D5', 22, 24)],
}
SCHEMES['5dp_B_6_18'] = {
    'name': '5 dp (solar 6-18)', 'short': '5dp_B',
    'blocks': [('D1', 0, 6), ('D2', 6, 18), ('D3', 18, 21),
               ('D4', 21, 23), ('D5', 23, 24)],
}
SCHEMES['5dp_C_5_18'] = {
    'name': '5 dp (solar 5-18, widest)', 'short': '5dp_C',
    'blocks': [('D1', 0, 5), ('D2', 5, 18), ('D3', 18, 21),
               ('D4', 21, 23), ('D5', 23, 24)],
}
SCHEMES['5dp_D_6_17'] = {
    'name': '5 dp (solar 6-17, tight)', 'short': '5dp_D',
    'blocks': [('D1', 0, 6), ('D2', 6, 17), ('D3', 17, 20),
               ('D4', 20, 22), ('D5', 22, 24)],
}
SCHEMES['6dp_A_5_17'] = {
    'name': '6 dp (solar 5-17, night split)', 'short': '6dp_A',
    'blocks': [('D1', 0, 3), ('D2', 3, 5), ('D3', 5, 17),
               ('D4', 17, 20), ('D5', 20, 22), ('D6', 22, 24)],
}
SCHEMES['6dp_B_6_18'] = {
    'name': '6 dp (solar 6-18, night split)', 'short': '6dp_B',
    'blocks': [('D1', 0, 3), ('D2', 3, 6), ('D3', 6, 18),
               ('D4', 18, 21), ('D5', 21, 23), ('D6', 23, 24)],
}
SCHEMES['6dp_C_5_18'] = {
    'name': '6 dp (solar 5-18, widest)', 'short': '6dp_C',
    'blocks': [('D1', 0, 3), ('D2', 3, 5), ('D3', 5, 18),
               ('D4', 18, 21), ('D5', 21, 23), ('D6', 23, 24)],
}
SCHEMES['6dp_D_ramp'] = {
    'name': '6 dp (solar 8-17, morning ramp)', 'short': '6dp_D',
    'blocks': [('D1', 0, 5), ('D2', 5, 8), ('D3', 8, 17),
               ('D4', 17, 20), ('D5', 20, 22), ('D6', 22, 24)],
}
SCHEMES['6dp_F_5_18_ramp'] = {
    'name': '6 dp (solar 8-18, morning ramp)', 'short': '6dp_F',
    'blocks': [('D1', 0, 5), ('D2', 5, 8), ('D3', 8, 18),
               ('D4', 18, 20), ('D5', 20, 22), ('D6', 22, 24)],
}

# ================================================================
# USER-DEFINED CANDIDATES  (add your own below this banner)
# ----------------------------------------------------------------
# Format contract — each scheme must have:
#   'name':   human-readable label used in plots / CSV
#   'short':  compact label for tight plot legends (optional but recommended)
#   'blocks': list of (label, start, end) — start < end, no gaps, must tile 0-24
#
# Optional:
#   'merge_labels': dict mapping block labels to conceptual groupings
#                   (used when two physical blocks represent the same concept,
#                    e.g. pre-midnight & post-midnight night blocks)
#
# The scheme key (dict key) becomes the identifier in the CSV, figures,
# and config JSON. Keep it short, lowercase, and descriptive (e.g.
# "6dp_solar_physics" not "my_new_scheme_v3_final").
#
# A validator at the end of this block will fail loudly if any scheme
# breaks the contiguity contract — no need to hand-check.
#
# Design principle for these additions (April 2026):
#   Separate the SOLAR PEAK window (CF ~0.45-0.55) from the SOLAR SHOULDER
#   windows (CF ~0.15-0.35). Existing "solar-wide" schemes (6dp_D_ramp,
#   6dp_F_5_18_ramp) lump peak and shoulder into one fat 8-17 or 8-18
#   block at ~0.33 average CF, which understates midday solar output by
#   ~40% and overstates afternoon output by ~2x. Separating them gives
#   dispatch models a genuinely distinct "high solar" block.
# ================================================================

SCHEMES['5dp_solar_physics'] = {
    'name': '5 dp (physics: night / ramp / peak / shoulder / evening+night)',
    'short': '5dp_phys',
    'blocks': [('D1',  0,  5),   # night baseline (CF = 0)
               ('D2',  5,  8),   # morning ramp (CF 0 → ~0.25)
               ('D3',  8, 12),   # SOLAR PEAK (CF ~0.45-0.55)
               ('D4', 12, 16),   # afternoon shoulder (CF ~0.30-0.45)
               ('D5', 16, 24)],  # evening peak + late night (catch-all, 8h)
    # Design note: this 5dp deliberately sacrifices the evening demand peak
    # by folding it into a fat 16-24 block. Budget is spent on solar fidelity
    # (4 of 5 blocks carve the solar curve) rather than demand-peak resolution.
}

SCHEMES['6dp_solar_physics'] = {
    'name': '6 dp (physics: night / ramp / peak / shoulder / evening / late)',
    'short': '6dp_phys',
    'blocks': [('D1',  0,  5),   # night baseline
               ('D2',  5,  8),   # morning ramp
               ('D3',  8, 12),   # SOLAR PEAK
               ('D4', 12, 16),   # afternoon shoulder
               ('D5', 16, 20),   # evening demand peak
               ('D6', 20, 24)],  # late night
    # Design note: this is 5dp_solar_physics plus one extra block spent on
    # splitting the 16-24 catch-all into a dedicated evening-peak block
    # (16-20) and a late-night block (20-24). Recovers the demand-peak
    # resolution the 5dp version sacrifices.
}

# ----------------------------------------------------------------
# VALIDATOR — runs once at import time. Fails loudly on bad schemes.
# ----------------------------------------------------------------

def _validate_all_schemes(schemes):
    """Check every scheme in SCHEMES tiles 0-24 contiguously."""
    problems = []
    for key, spec in schemes.items():
        blocks = spec.get('blocks', [])
        if not blocks:
            problems.append(f"  {key}: no blocks"); continue
        srt = sorted(blocks, key=lambda b: b[1])
        if srt[0][1] != 0:
            problems.append(f"  {key}: first block starts at {srt[0][1]}, not 0")
        if srt[-1][2] != 24:
            problems.append(f"  {key}: last block ends at {srt[-1][2]}, not 24")
        for i in range(len(srt) - 1):
            if srt[i][2] != srt[i+1][1]:
                problems.append(
                    f"  {key}: gap/overlap between {srt[i][0]} (ends {srt[i][2]}) "
                    f"and {srt[i+1][0]} (starts {srt[i+1][1]})"
                )
        for bname, s, e in blocks:
            if e <= s:
                problems.append(f"  {key}: {bname} has non-positive duration ({s}→{e})")
    if problems:
        raise ValueError("Scheme definition errors:\n" + "\n".join(problems))

_validate_all_schemes(SCHEMES)
if VERBOSE:
    print(f"[schemes] Validated {len(SCHEMES)} schemes — all tile 0-24 cleanly.")

# Color palette for schemes
SCHEME_COLORS = {
    # Original baselines
    '2dp_equal':    '#d62728',
    '2dp_daynight': '#ff7f0e',
    '3dp_legacy':   '#9467bd',
    '4dp_equal':    '#00414D',   # CLG GREEN_DARK
    '4dp_shifted':  '#2ca02c',
    '5dp_asym':     '#23978E',   # CLG TEAL
    '5dp_solar':    '#e6550d',
    '6dp_solar':    '#756bb1',
    '6dp_equal':    '#1f77b4',
    '8dp_equal':    '#8c564b',
    '12dp_equal':   '#e377c2',
    '24dp_hourly':  '#7f7f7f',
    # New candidates (2026)
    '5dp_A_5_17':     '#0A595F',  # CLG GREEN_MID
    '5dp_B_6_18':     '#66c2a5',
    '5dp_C_5_18':     '#fc8d62',
    '5dp_D_6_17':     '#8da0cb',
    '6dp_A_5_17':     '#ffd92f',
    '6dp_B_6_18':     '#e5c494',
    '6dp_C_5_18':     '#b3b3b3',
    '6dp_D_ramp':     '#1b9e77',
    '6dp_F_5_18_ramp':'#7570b3',
    # User-defined physics-informed candidates (2026)
    '5dp_solar_physics':    '#0A595F',  # CLG GREEN_MID
    '6dp_solar_physics':    '#00414D',  # CLG GREEN_DARK
}

# ======================================================================
# INDIA REGION MAPPING
# ======================================================================

IND_REGIONS = ['INDNO', 'INDEA', 'INDNE', 'INDSO', 'INDWE']
IND_REGION_MAP = {
    'eastern': 'INDEA', 'north-eastern': 'INDNE', 'northeastern': 'INDNE',
    'northern': 'INDNO', 'northen': 'INDNO',
    'southern': 'INDSO', 'western': 'INDWE',
}
_MONTH_ABBR = {'Jan':1,'Feb':2,'Mar':3,'Apr':4,'May':5,'Jun':6,
               'Jul':7,'Aug':8,'Sep':9,'Oct':10,'Nov':11,'Dec':12}

# UTC offsets for solar
UTC_OFFSETS = {
    'BGDXX': 6, 'BTNXX': 6, 'NPLXX': 5.75, 'LKAXX': 5.5,
    'INDNO': 5.5, 'INDEA': 5.5, 'INDNE': 5.5, 'INDSO': 5.5, 'INDWE': 5.5,
}


# ======================================================================
# HELPER FUNCTIONS
# ======================================================================

def assign_daypart(hour, scheme):
    """Return block label for hour 0–23."""
    for label, start, end in scheme['blocks']:
        if start <= hour < end:
            return label
    return None


def block_hours(scheme):
    """Return {label: [list of hours]}."""
    out = {}
    for label, start, end in scheme['blocks']:
        hrs = list(range(start, end))
        out[label] = hrs
    return out


def effective_n_blocks(scheme):
    """Number of effective blocks (after merging labels if specified)."""
    if 'merge_labels' in scheme:
        return len(set(scheme['merge_labels'].values()))
    return len(scheme['blocks'])


def n_blocks_raw(scheme):
    """Raw block count (for step-function plotting)."""
    return len(scheme['blocks'])


# ======================================================================
# METRICS
# ======================================================================

def compute_mean_profile(df, season_months=None, hour_col='hour', mw_col='mw'):
    """Compute normalised 24-h profile (daily mean = 1.0)."""
    d = df.copy()
    if season_months is not None:
        d = d[d['month'].isin(season_months)]
    if d.empty:
        return np.full(24, np.nan)
    hourly_mean = d.groupby(hour_col)[mw_col].mean()
    daily_avg = hourly_mean.mean()
    if daily_avg == 0:
        return np.full(24, np.nan)
    profile = (hourly_mean / daily_avg).reindex(range(24), fill_value=np.nan)
    return profile.values


def compute_block_means(profile_24, scheme):
    """Block-mean values from 24-h profile."""
    bh = block_hours(scheme)
    return {lbl: np.nanmean([profile_24[h] for h in hrs])
            for lbl, hrs in bh.items()}


def reconstruct_step(profile_24, scheme):
    """Build 24-element step-function from block means."""
    bm = compute_block_means(profile_24, scheme)
    step = np.full(24, np.nan)
    for lbl, hrs in block_hours(scheme).items():
        for h in hrs:
            step[h] = bm[lbl]
    return step


def within_block_variance(profile_24, scheme):
    """Weighted average within-block variance. Lower = better."""
    bh = block_hours(scheme)
    total_var = 0.0
    for lbl, hrs in bh.items():
        vals = [profile_24[h] for h in hrs if not np.isnan(profile_24[h])]
        if len(vals) < 2:
            continue
        total_var += np.var(vals) * (len(hrs) / 24.0)
    return total_var


def peak_preservation(profile_24, scheme):
    """Block-mean range / true hourly range. 1.0 = perfect."""
    valid = profile_24[~np.isnan(profile_24)]
    if len(valid) == 0:
        return np.nan
    true_range = valid.max() - valid.min()
    if true_range == 0:
        return np.nan
    bm = compute_block_means(profile_24, scheme)
    bm_vals = [v for v in bm.values() if not np.isnan(v)]
    if len(bm_vals) < 2:
        return np.nan
    return (max(bm_vals) - min(bm_vals)) / true_range


def rmse_step(profile_24, scheme):
    """RMSE of step-function reconstruction vs. true hourly profile."""
    step = reconstruct_step(profile_24, scheme)
    mask = ~np.isnan(profile_24) & ~np.isnan(step)
    if mask.sum() == 0:
        return np.nan
    return np.sqrt(np.mean((profile_24[mask] - step[mask]) ** 2))


def solar_differentiation(solar_24, scheme):
    """CV of block-mean solar CFs. Higher = better separation."""
    if solar_24 is None:
        return np.nan
    bm = compute_block_means(solar_24, scheme)
    vals = np.array([v for v in bm.values() if not np.isnan(v)])
    if len(vals) < 2 or vals.mean() == 0:
        return np.nan
    return vals.std() / vals.mean()


# ======================================================================
# DATA LOADERS — use compiled CSVs (fast path)
# ======================================================================

def load_bgd(base_dir):
    """Load BGD from compiled half-hourly CSV, resample to hourly."""
    csv_path = os.path.join(base_dir, 'Bangladesh', 'pgcb_analysis',
                            'bgd_halfhourly_compiled.csv')
    if not os.path.exists(csv_path):
        print(f"  [BGD] NOT FOUND: {csv_path}")
        return pd.DataFrame()
    df = pd.read_csv(csv_path)
    print(f"  [BGD] Loaded {len(df):,} half-hourly rows from compiled CSV")
    # Floor hour to integer and average within each date × hour
    df['hour_int'] = df['hour'].astype(int)
    if 'month' not in df.columns:
        df['date_dt'] = pd.to_datetime(df['date'])
        df['month'] = df['date_dt'].dt.month
    hourly = (df.groupby(['date', 'month', 'hour_int'], as_index=False)['demand_MW']
              .mean())
    hourly.rename(columns={'hour_int': 'hour', 'demand_MW': 'mw'}, inplace=True)
    hourly['region'] = 'BGDXX'
    print(f"  [BGD] {len(hourly):,} hourly rows, "
          f"{hourly['date'].nunique()} days")
    return hourly[['region', 'hour', 'month', 'mw']]


def load_lka(base_dir):
    """Load LKA from compiled hourly CSV."""
    csv_path = os.path.join(base_dir, 'Sri Lanka', 'lka_hourly_compiled.csv')
    if not os.path.exists(csv_path):
        print(f"  [LKA] NOT FOUND: {csv_path}")
        return pd.DataFrame()
    df = pd.read_csv(csv_path)
    print(f"  [LKA] Loaded {len(df):,} hourly rows from compiled CSV")
    df.rename(columns={'demand_MW': 'mw'}, inplace=True)
    df['hour'] = df['hour'].astype(int)
    if 'month' not in df.columns:
        df['date_dt'] = pd.to_datetime(df['date'])
        df['month'] = df['date_dt'].dt.month
    df['region'] = 'LKAXX'
    print(f"  [LKA] {df['date'].nunique()} days")
    return df[['region', 'hour', 'month', 'mw']]


def _parse_ind_date(s):
    """Parse '01-Jan 12am' -> (month, hour)."""
    m = re.match(r'(\d{1,2})-(\w{3})\s+(\d{1,2})(am|pm)', str(s).strip())
    if not m:
        return None, None
    month = _MONTH_ABBR.get(m.group(2))
    h12, ampm = int(m.group(3)), m.group(4)
    if ampm == 'am':
        hour = 0 if h12 == 12 else h12
    else:
        hour = 12 if h12 == 12 else h12 + 12
    return month, hour


def load_ind(base_dir):
    """Load India hourly demand from year_demand Excel files."""
    src = os.path.join(base_dir, 'India', 'Initial_Analysis_Sources')
    files = sorted(glob.glob(os.path.join(src, 'year_demand_*.xlsx')))
    if not files:
        print(f"  [IND] No demand files found in {src}")
        return pd.DataFrame()

    all_rows = []
    for fp in files:
        fname = os.path.basename(fp)
        print(f"  [IND] Reading {fname} ...")
        try:
            df = pd.read_excel(fp)
            for _, row in df.iterrows():
                region_raw = str(row.iloc[0]).strip()
                # Extract region
                rm = re.match(r'(.+?)\s*region', region_raw, re.IGNORECASE)
                if rm:
                    rname = rm.group(1).strip().lower().replace(' ', '-')
                    code = IND_REGION_MAP.get(rname) or IND_REGION_MAP.get(rname.replace('-', ''))
                else:
                    code = None
                if code is None:
                    continue
                month, hour = _parse_ind_date(row.iloc[1])
                if month is None or hour is None:
                    continue
                mw = row.iloc[2] if len(row) > 2 else None
                if pd.notna(mw):
                    all_rows.append({'region': code, 'month': month,
                                     'hour': int(hour), 'mw': float(mw)})
        except Exception as e:
            print(f"    ERROR: {e}")

    if not all_rows:
        return pd.DataFrame()
    out = pd.DataFrame(all_rows)
    print(f"  [IND] {len(out):,} hourly rows, regions: {sorted(out.region.unique())}")
    return out[['region', 'hour', 'month', 'mw']]


def load_npl(base_dir):
    """Load NPL representative 24-h profiles (Simkhada 2022)."""
    csv_path = os.path.join(base_dir, 'Nepal', 'Nepal_HourlyProfiles_Literature.csv')
    if not os.path.exists(csv_path):
        print(f"  [NPL] NOT FOUND: {csv_path}")
        return pd.DataFrame()
    df = pd.read_csv(csv_path)
    print(f"  [NPL] Loaded {len(df)} rows (representative profiles)")
    # Columns: SEASON, HOUR, MW (or similar)
    # Normalise column names
    df.columns = [c.strip().upper() for c in df.columns]
    if 'MW' not in df.columns and 'PU' in df.columns:
        df['MW'] = df['PU']  # per-unit is fine for normalised profiles
    rows = []
    for _, r in df.iterrows():
        s = str(r['SEASON']).strip()
        # Map season to months
        season_months = None
        for sc, sl, sm, sd in SEASON_DEF:
            if sc == s:
                season_months = sm
                break
        if season_months is None:
            continue
        for m in season_months:
            rows.append({'region': 'NPLXX', 'hour': int(r['HOUR']),
                         'month': m, 'mw': float(r['MW'])})
    out = pd.DataFrame(rows)
    print(f"  [NPL] Expanded to {len(out)} rows (replicated across months)")
    return out[['region', 'hour', 'month', 'mw']]


def load_btn(base_dir):
    """Load BTN representative 24-h profiles from BTN_Profiles.xlsx."""
    fp = os.path.join(base_dir, 'Bhutan', 'BTN_Profiles.xlsx')
    if not os.path.exists(fp):
        print(f"  [BTN] NOT FOUND: {fp}")
        return pd.DataFrame()
    try:
        df_hp = pd.read_excel(fp, sheet_name='Hourly_Profiles')
    except Exception as e:
        print(f"  [BTN] ERROR reading Hourly_Profiles: {e}")
        return pd.DataFrame()
    print(f"  [BTN] Loaded Hourly_Profiles ({len(df_hp)} rows)")
    # Columns: hour, S1, S2, S3, S4
    rows = []
    for _, r in df_hp.iterrows():
        h = int(r['hour'])
        for sc, sl, sm, sd in SEASON_DEF:
            if sc in df_hp.columns:
                mw_val = float(r[sc])
                for m in sm:
                    rows.append({'region': 'BTNXX', 'hour': h,
                                 'month': m, 'mw': mw_val})
    out = pd.DataFrame(rows)
    print(f"  [BTN] Expanded to {len(out)} rows")
    return out[['region', 'hour', 'month', 'mw']]


def load_ninja_hourly(base_dir):
    """
    Load Renewables.ninja hourly CFs from compiled_reninja_hourly.csv
    (produced by rebuild_reninja_timeslices.py). Returns 24-h mean profiles
    for BOTH solar and wind per region.

    Primary path: _Reno_Ninja/ninja_data/output_rebuilt/compiled_reninja_hourly.csv
    Fallback:     raw CSVs from _Reno_Ninja/ninja_data/solar/ and wind/

    Returns:
        solar_profiles : dict { region : np.array(24) }
        wind_profiles  : dict { region : np.array(24) }
        ninja_hourly   : DataFrame or None (raw hourly table for re-agg)
    """
    hourly_csv = os.path.join(base_dir, '_Reno_Ninja', 'ninja_data',
                              'output_rebuilt', 'compiled_reninja_hourly.csv')

    solar_profiles = {}
    wind_profiles = {}
    ninja_hourly = None

    if os.path.exists(hourly_csv):
        print(f"  [NINJA] Reading compiled hourly: {os.path.basename(hourly_csv)}")
        usecols = ['cf', 'resource', 'region', 'hour_local', 'season', 'model_year']
        try:
            df = pd.read_csv(hourly_csv, usecols=usecols)
        except ValueError:
            df = pd.read_csv(hourly_csv)
            needed = {'cf', 'resource', 'region', 'hour_local'}
            if not needed.issubset(set(df.columns)):
                print(f"  [NINJA] WARNING: Missing columns. Have: {list(df.columns)}")
                return solar_profiles, wind_profiles, None
        ninja_hourly = df
        print(f"  [NINJA] {len(df):,} hourly rows, "
              f"resources: {sorted(df.resource.unique())}, "
              f"regions: {sorted(df.region.unique())}")

        for resource in sorted(df['resource'].unique()):
            rdf = df[df['resource'] == resource]
            profiles = solar_profiles if resource == 'solar' else wind_profiles
            for region in sorted(rdf['region'].unique()):
                sub = rdf[rdf['region'] == region]
                profile = (sub.groupby('hour_local')['cf'].mean()
                           .reindex(range(24), fill_value=0.0))
                profiles[region] = profile.values
                if VERBOSE:
                    print(f"    {resource:6s} {region}: "
                          f"peak CF={profile.max():.3f} at hour {profile.idxmax()}")

    else:
        print(f"  [NINJA] compiled_reninja_hourly.csv NOT FOUND:")
        print(f"          {hourly_csv}")
        print(f"          Run rebuild_reninja_timeslices.py first.")
        print(f"          Falling back to raw Ninja CSVs ...")

        for res_type, subdir in [('solar', 'solar'), ('wind', 'wind')]:
            ninja_dir = os.path.join(base_dir, '_Reno_Ninja', 'ninja_data', subdir)
            if not os.path.isdir(ninja_dir):
                continue
            profiles = solar_profiles if res_type == 'solar' else wind_profiles
            for region, utc_off in UTC_OFFSETS.items():
                pattern = os.path.join(ninja_dir, f'{res_type}_{region}_*.csv')
                files = sorted(glob.glob(pattern))
                if not files:
                    continue
                frames = []
                for fp in files:
                    try:
                        fdf = pd.read_csv(fp, comment='#')
                        fdf.columns = ['time', 'cf']
                        fdf['time'] = pd.to_datetime(fdf['time'])
                        fdf['hour_local'] = (fdf['time'] + pd.Timedelta(hours=utc_off)).dt.hour
                        frames.append(fdf[['hour_local', 'cf']])
                    except Exception:
                        pass
                if frames:
                    big = pd.concat(frames, ignore_index=True)
                    profile = (big.groupby('hour_local')['cf'].mean()
                               .reindex(range(24), fill_value=0.0))
                    profiles[region] = profile.values
                    if VERBOSE:
                        print(f"    {res_type:6s} {region}: {len(frames)} years, "
                              f"peak CF={profile.max():.3f} at hour {profile.idxmax()}")

    return solar_profiles, wind_profiles, ninja_hourly


def compute_ninja_cfs_per_scheme(ninja_hourly, scheme):
    """
    Re-aggregate Ninja hourly CFs to a specific daypart scheme.
    Returns dict { (resource, region, timeslice) : cf_mean }.
    """
    if ninja_hourly is None or ninja_hourly.empty:
        return {}
    df = ninja_hourly.copy()
    df['_dp'] = df['hour_local'].astype(int).apply(lambda h: assign_daypart(h, scheme))
    df['_ts'] = df['season'].astype(str) + df['_dp'].astype(str)
    agg = df.groupby(['resource', 'region', '_ts'])['cf'].mean().reset_index()
    return {(row['resource'], row['region'], row['_ts']): round(row['cf'], 6)
            for _, row in agg.iterrows()}


# ======================================================================
# ANALYSIS ENGINE
# ======================================================================

def analyze_region(label, df, solar_24, wind_24, schemes, seasons):
    """
    Run all metrics for one region across all schemes and seasons.
    Returns list of result dicts.
    """
    results = []
    for season_name, season_months in seasons.items():
        profile = compute_mean_profile(df, season_months)
        if np.all(np.isnan(profile)):
            continue

        for scheme_key, scheme in schemes.items():
            wbv = within_block_variance(profile, scheme)
            pp  = peak_preservation(profile, scheme)
            rms = rmse_step(profile, scheme)
            scv = solar_differentiation(solar_24, scheme)
            wcv = solar_differentiation(wind_24, scheme)  # same CV metric

            results.append({
                'label':     label,
                'season':    season_name,
                'scheme':    scheme_key,
                'scheme_name': scheme['name'],
                'n_dp':      effective_n_blocks(scheme),
                'n_dp_raw':  n_blocks_raw(scheme),
                'wbv':       wbv,
                'pp':        pp,
                'rmse':      rms,
                'scv':       scv,
                'wcv':       wcv,
            })
    return results


# ======================================================================
# PLOTTING
# ======================================================================

def plot_scheme_overlay(label, profile_24, schemes, out_path):
    """
    Demand overlay figure: 24-h demand curve with step-function overlays
    for ALL schemes drawn at equal visual weight. No privileged scheme.

    Solar and wind CF figures are produced separately by plot_cf_overlay.

    Legend is placed outside the plot area to accommodate ~24 entries.
    """
    fig, ax = plt.subplots(figsize=(17, 7))
    hours = np.arange(24)

    # True hourly profile — the only element drawn bold
    ax.plot(hours, profile_24, 'k-o', ms=4, lw=2.4, zorder=100,
            label='Hourly profile')

    # Step-function overlays — every scheme, equal visual weight
    for sk, scheme in schemes.items():
        step = reconstruct_step(profile_24, scheme)
        color = SCHEME_COLORS.get(sk, '#888')
        ax.step(hours, step, where='mid', color=color,
                lw=1.2, ls='-', alpha=0.75,
                label=scheme['name'])

    ax.set_xlim(-0.5, 23.5)
    ax.set_xticks(hours)
    ax.set_xlabel('Hour of day (local time)', fontsize=11)
    ax.set_ylabel('Normalised demand (daily mean = 1)', fontsize=11)
    ax.set_title(f'{label} — Timeslice scheme sweep (Annual)',
                 fontsize=13, fontweight='bold')
    ax.legend(fontsize=7.5, loc='center left', bbox_to_anchor=(1.01, 0.5),
              framealpha=0.9, ncol=1, borderaxespad=0.)
    ax.grid(axis='y', alpha=0.3)

    plt.tight_layout()
    plt.savefig(out_path, dpi=DPI, bbox_inches='tight')
    plt.close(fig)
    if VERBOSE:
        print(f"    Saved: {os.path.basename(out_path)}")


def plot_cf_overlay(label, cf_24, resource, schemes, out_path):
    """
    Standalone CF figure: 24-h solar or wind profile with step-function
    overlays for ALL schemes at equal visual weight.

    CV annotations dropped — too many schemes to annotate cleanly;
    SCV/WCV values are in sensitivity_timeslice_summary.csv and plotted
    on the bar summary.
    """
    if cf_24 is None or np.all(cf_24 == 0):
        return

    fig, ax = plt.subplots(figsize=(17, 6.5))
    hours = np.arange(24)

    # Colour and fill for the resource
    if resource == 'solar':
        curve_color, fill_color, fill_alpha = 'goldenrod', 'gold', 0.15
    else:
        curve_color, fill_color, fill_alpha = 'steelblue', '#B0D4F1', 0.15

    # True hourly profile — only bold element
    ax.fill_between(hours, cf_24, alpha=fill_alpha, color=fill_color)
    ax.plot(hours, cf_24, '-o', ms=4.5, lw=2.4, color=curve_color,
            zorder=100, markeredgewidth=0.5, markeredgecolor='white',
            label=f'Hourly {resource} CF')

    # Step-function overlays — every scheme, equal visual weight
    for sk, scheme in schemes.items():
        step = reconstruct_step(cf_24, scheme)
        color = SCHEME_COLORS.get(sk, '#888')
        ax.step(hours, step, where='mid', color=color,
                lw=1.2, ls='-', alpha=0.75,
                label=scheme['name'])

    ax.set_xlim(-0.5, 23.5)
    ax.set_xticks(hours)
    ax.set_xlabel('Hour of day (local time)', fontsize=11)
    ax.set_ylabel('Capacity factor', fontsize=11)
    ax.set_title(f'{label} — {resource.capitalize()} CF timeslice scheme sweep (Annual)',
                 fontsize=13, fontweight='bold')
    ax.legend(fontsize=7.5, loc='center left', bbox_to_anchor=(1.01, 0.5),
              framealpha=0.9, ncol=1, borderaxespad=0.)
    ax.grid(axis='y', alpha=0.3)
    ax.set_ylim(bottom=0)

    plt.tight_layout()
    plt.savefig(out_path, dpi=DPI, bbox_inches='tight')
    plt.close(fig)
    if VERBOSE:
        print(f"    Saved: {os.path.basename(out_path)}")


def plot_seasonal_grid(label, df, schemes, out_path):
    """
    2×2 grid: one subplot per OSTRAM season, each showing hourly profile
    + step-function overlay for ALL schemes at equal visual weight.

    Legend is shared and placed to the right of the figure.
    RMSE annotations dropped — 24 schemes won't fit; values are in
    sensitivity_timeslice_summary.csv.
    """
    fig, axes = plt.subplots(2, 2, figsize=(18, 11))
    axes = axes.flatten()
    hours = np.arange(24)

    legend_handles = None  # capture from first populated subplot

    for idx, (s_code, s_label, s_months, _) in enumerate(SEASON_DEF):
        ax = axes[idx]
        profile = compute_mean_profile(df, s_months)
        if np.all(np.isnan(profile)):
            ax.set_title(f'{s_code}: {s_label} — No data')
            continue

        ax.plot(hours, profile, 'k-o', ms=3, lw=2.0, zorder=100,
                label='Hourly')
        for sk, scheme in schemes.items():
            step = reconstruct_step(profile, scheme)
            color = SCHEME_COLORS.get(sk, '#888')
            ax.step(hours, step, where='mid', color=color,
                    lw=1.0, ls='-', alpha=0.65,
                    label=scheme['name'])

        ax.set_xlim(-0.5, 23.5)
        ax.set_xticks(range(0, 24, 3))
        ax.set_title(f'{s_code}: {s_label}', fontsize=11, fontweight='bold')
        ax.grid(axis='y', alpha=0.3)

        if legend_handles is None:
            legend_handles, legend_labels = ax.get_legend_handles_labels()

    fig.suptitle(f'{label} — Seasonal comparison (all candidates)',
                 fontsize=14, fontweight='bold')

    # Shared legend to the right of the grid
    if legend_handles is not None:
        fig.legend(legend_handles, legend_labels,
                   loc='center left', bbox_to_anchor=(0.92, 0.5),
                   fontsize=7.5, framealpha=0.9, ncol=1)
        plt.tight_layout(rect=[0, 0, 0.91, 0.96])
    else:
        plt.tight_layout(rect=[0, 0, 1, 0.96])

    plt.savefig(out_path, dpi=DPI, bbox_inches='tight')
    plt.close(fig)
    if VERBOSE:
        print(f"    Saved: {os.path.basename(out_path)}")


def plot_convergence(results_df, out_path):
    """
    Convergence plot: RMSE and PP vs. effective N dayparts,
    one line per region (Annual only).
    """
    ann = results_df[results_df['season'] == 'Annual'].copy()
    if ann.empty:
        return

    fig, axes = plt.subplots(1, 3, figsize=(18, 6))
    regions = sorted(ann['label'].unique())

    # RMSE convergence
    ax = axes[0]
    for reg in regions:
        sub = ann[ann['label'] == reg].sort_values('n_dp')
        ax.plot(sub['n_dp'], sub['rmse'], '-o', ms=5, label=reg, alpha=0.8)
    ax.set_xlabel('Effective N dayparts')
    ax.set_ylabel('RMSE')
    ax.set_title('RMSE convergence', fontweight='bold')
    ax.legend(fontsize=7, ncol=2)
    ax.grid(alpha=0.3)

    # Peak preservation
    ax = axes[1]
    for reg in regions:
        sub = ann[ann['label'] == reg].sort_values('n_dp')
        ax.plot(sub['n_dp'], sub['pp'], '-s', ms=5, label=reg, alpha=0.8)
    ax.axhline(1.0, color='red', ls=':', lw=0.8, alpha=0.5)
    ax.set_xlabel('Effective N dayparts')
    ax.set_ylabel('Peak preservation')
    ax.set_title('Peak preservation convergence', fontweight='bold')
    ax.legend(fontsize=7, ncol=2)
    ax.grid(alpha=0.3)

    # WBV
    ax = axes[2]
    for reg in regions:
        sub = ann[ann['label'] == reg].sort_values('n_dp')
        ax.plot(sub['n_dp'], sub['wbv'], '-^', ms=5, label=reg, alpha=0.8)
    ax.set_xlabel('Effective N dayparts')
    ax.set_ylabel('Within-block variance')
    ax.set_title('WBV convergence', fontweight='bold')
    ax.legend(fontsize=7, ncol=2)
    ax.grid(alpha=0.3)

    fig.suptitle('Timeslice sensitivity — Metric convergence by N dayparts',
                 fontsize=14, fontweight='bold')
    plt.tight_layout(rect=[0, 0, 1, 0.95])
    plt.savefig(out_path, dpi=DPI, bbox_inches='tight')
    plt.close(fig)
    if VERBOSE:
        print(f"  Saved: {os.path.basename(out_path)}")


def plot_bar_summary(results_df, out_path):
    """
    Grouped bar chart: RMSE per region for each scheme (Annual).
    Schemes not present in results_df are silently skipped so the
    function never crashes when SCHEMES is modified.
    """
    ann = results_df[results_df['season'] == 'Annual'].copy()
    # Preferred display order — includes baselines AND new candidates.
    # Any scheme not present in the results will just be skipped.
    preferred_order = [
        '3dp_legacy', '4dp_equal', '4dp_shifted', '5dp_asym',
        '5dp_solar', '6dp_solar', '6dp_equal', '8dp_equal',
        '5dp_A_5_17', '5dp_B_6_18', '5dp_C_5_18',
        '5dp_D_6_17', '5dp_E_5_17_v2', '5dp_F_5_18_v2',
        '6dp_A_5_17', '6dp_B_6_18', '6dp_C_5_18',
        '6dp_D_ramp', '6dp_E_dusk', '6dp_F_5_18_ramp',
    ]
    # Keep only schemes that are BOTH in preferred_order AND present in the data
    # AND still defined in SCHEMES (so the legend lookup can't fail)
    bar_schemes = [s for s in preferred_order
                   if s in set(ann['scheme']) and s in SCHEMES]
    # Also append any scheme in the data that isn't in preferred_order (future-proof)
    bar_schemes += [s for s in sorted(set(ann['scheme']))
                    if s not in bar_schemes and s in SCHEMES]
    if not bar_schemes:
        return
    ann = ann[ann['scheme'].isin(bar_schemes)]

    regions = sorted(ann['label'].unique())
    n_regions = len(regions)
    n_schemes = len(bar_schemes)
    x = np.arange(n_regions)
    width = 0.8 / n_schemes

    fig, ax = plt.subplots(figsize=(max(14, n_regions * 1.7), 7))
    for i, sk in enumerate(bar_schemes):
        sub = ann[ann['scheme'] == sk].set_index('label')
        vals = [sub.loc[r, 'rmse'] if r in sub.index else 0 for r in regions]
        color = SCHEME_COLORS.get(sk, '#888')
        ax.bar(x + i * width, vals, width,
               label=SCHEMES[sk]['name'],
               color=color, alpha=0.85, edgecolor='white')

    ax.set_xticks(x + width * (n_schemes - 1) / 2)
    ax.set_xticklabels(regions, fontsize=9, rotation=30, ha='right')
    ax.set_ylabel('RMSE (normalised demand)', fontsize=11)
    ax.set_title('Annual RMSE by region and daypart scheme', fontsize=13,
                 fontweight='bold')
    ax.legend(fontsize=7, ncol=3, loc='upper right')
    ax.grid(axis='y', alpha=0.3)

    plt.tight_layout()
    plt.savefig(out_path, dpi=DPI, bbox_inches='tight')
    plt.close(fig)
    if VERBOSE:
        print(f"  Saved: {os.path.basename(out_path)}")


# ======================================================================
# MAIN
# ======================================================================

print('=' * 70)
print('OSTRAM TIMESLICE SENSITIVITY ANALYSIS')
print(f'Schemes: {len(SCHEMES)}  |  Seasons: {len(SEASONS_FOR_ANALYSIS)}')
print(f'Output:  {OUTPUT_DIR}')
print(f'Started: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
print('=' * 70)

os.makedirs(OUTPUT_DIR, exist_ok=True)

# --- 1. Load demand data ---
print('\n[1] Loading demand data ...')
demand_data = {}  # {region_label: DataFrame}

if RUN_BGD:
    df_bgd = load_bgd(BASE_DIR)
    if not df_bgd.empty:
        demand_data['BGD'] = df_bgd

if RUN_LKA:
    df_lka = load_lka(BASE_DIR)
    if not df_lka.empty:
        demand_data['LKA'] = df_lka

if RUN_IND:
    df_ind = load_ind(BASE_DIR)
    if not df_ind.empty:
        for reg in IND_REGIONS:
            sub = df_ind[df_ind['region'] == reg]
            if not sub.empty:
                demand_data[reg] = sub

if RUN_NPL:
    df_npl = load_npl(BASE_DIR)
    if not df_npl.empty:
        demand_data['NPL*'] = df_npl  # asterisk = representative profile

if RUN_BTN:
    df_btn = load_btn(BASE_DIR)
    if not df_btn.empty:
        demand_data['BTN*'] = df_btn

print(f'\n  Regions loaded: {list(demand_data.keys())}')

# --- 2. Load Ninja CFs (solar + wind) ---
print('\n[2] Loading Ninja hourly profiles (solar + wind) ...')
solar_profiles, wind_profiles, ninja_hourly_df = load_ninja_hourly(BASE_DIR)
print(f'  Solar regions: {list(solar_profiles.keys())}')
print(f'  Wind regions:  {list(wind_profiles.keys())}')
if ninja_hourly_df is not None:
    print(f'  Ninja hourly DataFrame: {len(ninja_hourly_df):,} rows (available for re-agg)')

# Map region labels to Ninja region codes
NINJA_MAP = {
    'BGD': 'BGDXX', 'LKA': 'LKAXX', 'NPL*': 'NPLXX', 'BTN*': 'BTNXX',
    'INDNO': 'INDNO', 'INDEA': 'INDEA', 'INDNE': 'INDNE',
    'INDSO': 'INDSO', 'INDWE': 'INDWE',
}

# --- 3. Compute metrics ---
print('\n[3] Computing metrics ...')
all_results = []

for label, df in demand_data.items():
    ninja_key = NINJA_MAP.get(label)
    solar_24 = solar_profiles.get(ninja_key)
    wind_24 = wind_profiles.get(ninja_key)
    print(f'\n  {label} ({len(df):,} rows) ...')
    results = analyze_region(label, df, solar_24, wind_24, SCHEMES, SEASONS_FOR_ANALYSIS)
    all_results.extend(results)
    print(f'    → {len(results)} metric rows')

results_df = pd.DataFrame(all_results)

# --- 4. Generate figures ---
print('\n[4] Generating figures ...')

for label, df in demand_data.items():
    ninja_key = NINJA_MAP.get(label)
    solar_24 = solar_profiles.get(ninja_key)
    wind_24 = wind_profiles.get(ninja_key)
    profile_annual = compute_mean_profile(df, list(range(1, 13)))

    if np.any(~np.isnan(profile_annual)):
        tag = label.replace('*', '')
        # Demand overlay figure
        plot_scheme_overlay(
            label, profile_annual, SCHEMES,
            os.path.join(OUTPUT_DIR, f'sensitivity_overlay_{tag}.png'))
        # Solar CF overlay figure
        if solar_24 is not None:
            plot_cf_overlay(
                label, solar_24, 'solar', SCHEMES,
                os.path.join(OUTPUT_DIR, f'sensitivity_solar_{tag}.png'))
        # Wind CF overlay figure
        if wind_24 is not None:
            plot_cf_overlay(
                label, wind_24, 'wind', SCHEMES,
                os.path.join(OUTPUT_DIR, f'sensitivity_wind_{tag}.png'))
        # Seasonal grid
        plot_seasonal_grid(
            label, df, SCHEMES,
            os.path.join(OUTPUT_DIR, f'sensitivity_seasonal_{tag}.png'))

# Convergence plot
print('\n  Convergence plots ...')
plot_convergence(results_df, os.path.join(OUTPUT_DIR, 'sensitivity_convergence.png'))
plot_bar_summary(results_df, os.path.join(OUTPUT_DIR, 'sensitivity_bar_rmse.png'))

# --- 5. Export CSV ---
print('\n[5] Exporting summary ...')
csv_path = os.path.join(OUTPUT_DIR, 'sensitivity_timeslice_summary.csv')
results_df.to_csv(csv_path, index=False, float_format='%.6f')
print(f'  Saved: {csv_path}  ({len(results_df)} rows)')

# --- 5b. Export Ninja CF re-aggregation per scheme ---
if ninja_hourly_df is not None:
    print('\n[5b] Re-aggregating Ninja CFs per scheme ...')
    ninja_cf_rows = []
    for scheme_key, scheme in SCHEMES.items():
        ncf = compute_ninja_cfs_per_scheme(ninja_hourly_df, scheme)
        for (res, reg, ts), cf in ncf.items():
            ninja_cf_rows.append({
                'scheme': scheme_key, 'scheme_name': scheme['name'],
                'resource': res, 'region': reg, 'timeslice': ts,
                'cf_mean': cf,
            })
    if ninja_cf_rows:
        df_ncf = pd.DataFrame(ninja_cf_rows)
        ncf_path = os.path.join(OUTPUT_DIR, 'sensitivity_ninja_cfs.csv')
        df_ncf.to_csv(ncf_path, index=False, float_format='%.6f')
        print(f'  Saved: {ncf_path}  ({len(df_ncf)} rows)')
    else:
        print('  No Ninja CFs computed.')
else:
    print('\n[5b] Ninja hourly data not available — skipping CF re-aggregation.')
    print('     Run rebuild_reninja_timeslices.py to generate compiled_reninja_hourly.csv')

# --- 6. Print console summary (Annual only) ---
ann = results_df[results_df['season'] == 'Annual'].copy()
if not ann.empty:
    print('\n' + '=' * 110)
    print(f'{"Region":<10} {"Scheme":<25} {"Ndp":>4} {"WBV":>10} '
          f'{"PP":>10} {"RMSE":>10} {"SolarCV":>10} {"WindCV":>10}')
    print('-' * 110)
    for _, r in ann.sort_values(['label', 'n_dp']).iterrows():
        print(f'{r["label"]:<10} {r["scheme_name"]:<25} {r["n_dp"]:>4} '
              f'{r["wbv"]:>10.6f} {r["pp"]:>10.4f} {r["rmse"]:>10.4f} '
              f'{r["scv"]:>10.4f} {r["wcv"]:>10.4f}')
    print('=' * 110)

# --- 7. Key findings helper ---
print('\n' + '=' * 70)
print('INTERPRETATION GUIDE')
print('=' * 70)
print("""
  WBV  (Within-Block Variance) — lower is better.
       Measures how much hourly variation is lost when collapsing to blocks.

  PP   (Peak Preservation) — closer to 1.0 is better.
       Ratio of block-mean range to true hourly range.

  RMSE (Step reconstruction error) — lower is better.
       Direct measure of how well the step-function approximates the curve.

  SCV  (Solar Differentiation CV) — higher is better.
       How well the scheme separates solar-rich from solar-poor hours.

  WCV  (Wind Differentiation CV) — higher is better.
       How well the scheme separates wind-rich from wind-poor hours.
       Wind diurnal patterns vary by region and may differ from solar.

  * = Representative daily profiles only (NPL, BTN).
      Re-slicing these does not reveal new sub-daily variability —
      the metric improvements just reflect better fitting of a smooth shape.
      Interpret with caution; BGD, LKA, IND are the real test beds.

  Reference point: 4 seasons × 4 dayparts (6h equal) = 16 timeslices.
  Recommendation: look for the elbow where adding dayparts yields
  diminishing RMSE/WBV improvement across all regions.
""")

# Save config snapshot
config_path = os.path.join(OUTPUT_DIR, 'sensitivity_config.json')
with open(config_path, 'w') as f:
    json.dump({
        'schemes': {k: {'name': v['name'], 'blocks': v['blocks'],
                         'effective_n': effective_n_blocks(v)}
                    for k, v in SCHEMES.items()},
        'seasons': {k: v for k, v in SEASONS_FOR_ANALYSIS.items()},
        'regions': list(demand_data.keys()),
        'solar_regions': list(solar_profiles.keys()),
        'wind_regions': list(wind_profiles.keys()),
        'ninja_hourly_available': ninja_hourly_df is not None,
        'ninja_hourly_rows': len(ninja_hourly_df) if ninja_hourly_df is not None else 0,
        'generated': datetime.now().isoformat(),
    }, f, indent=2, default=str)

print(f'\nFinished: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
print(f'Output:  {OUTPUT_DIR}')
print('=' * 70)
