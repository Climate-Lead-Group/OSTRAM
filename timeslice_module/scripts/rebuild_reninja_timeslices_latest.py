# -*- coding: utf-8 -*-
"""
build_reninja_timeslices.py
==========================
Consolidates Renewables.ninja hourly CSV exports (solar PV and wind)
into a single long-format CSV with OSeMOSYS timeslice mapping.

Timeslice structure is FULLY CONFIGURABLE via DAYPART_DEF and SEASON_DEF
in the USER CONFIGURATION block. Change the number, boundaries, or labels
of dayparts/seasons in one place — everything downstream derives automatically.

UTC handling:
  Renewables.ninja timestamps are UTC. Daypart assignment requires local
  time. The script auto-detects UTC offset from param_lon (round(lon/15))
  or uses manual overrides in UTC_OFFSETS dict.

December assignment:
  December of calendar year Y is assigned to S1 of model_year Y+1.

Outputs
-------
1. compiled_reninja_hourly.csv  — hourly long-format with timeslice columns
2. compiled_reninja_ts.csv      — timeslice-aggregated mean CFs
                                   (resource × region × model_year × timeslice)
                                   with YearSplit — ready for OSeMOSYS
3. catalog_reninja.csv          — one row per source file with metadata
4. timeslice_config.json        — snapshot of the daypart/season config used
                                   (so downstream scripts can verify alignment)

Open in Spyder and press F5.

Author: CLG / Luis Victor-Gallardo
"""

import os
import json
import glob
import pandas as pd
import numpy as np
import time

# ╔══════════════════════════════════════════════════════════════════════╗
# ║  USER CONFIGURATION — edit these before running                    ║
# ╚══════════════════════════════════════════════════════════════════════╝

SOLAR_DIR  = "./ninja_data/solar"
WIND_DIR   = "./ninja_data/wind"
OUTPUT_DIR = "./ninja_data/output_rebuilt"

# Manual UTC offset overrides by region code (hours east of UTC).
# If a region is not listed here, offset is auto-detected from longitude.
UTC_OFFSETS = {
    'INDNO': 5.5, 'INDEA': 5.5, 'INDNE': 5.5, 'INDSO': 5.5, 'INDWE': 5.5,
    'NPLXX': 5.75, 'LKAXX': 5.5, 'BGDXX': 6, 'BTNXX': 6, 'MDVXX': 5,
}

# ╔══════════════════════════════════════════════════════════════════════╗
# ║  TIMESLICE SYSTEM DEFINITION                                       ║
# ║                                                                    ║
# ║  Edit SEASON_DEF and DAYPART_DEF to reshape the timeslice grid.    ║
# ║  Everything else (TIMESLICES, YEARSPLIT, hour_to_daypart, etc.)    ║
# ║  is derived automatically — do NOT edit below the line.            ║
# ╚══════════════════════════════════════════════════════════════════════╝

# --- SEASON DEFINITION ---
# Each entry: (code, label, list_of_months, days_in_season)
# Rules:
#   - Every month 1–12 must appear exactly once across all seasons
#   - days_in_season must sum to 365 (non-leap convention)
#   - December (month 12) → model_year + 1  (hardcoded convention)
SEASON_DEF = [
    ('S1', 'Winter (Dec-Feb)',       [12, 1, 2],       90),
    ('S2', 'Pre-monsoon (Mar-May)',  [3, 4, 5],        92),
    ('S3', 'SW Monsoon (Jun-Sep)',   [6, 7, 8, 9],    122),
    ('S4', 'Post-monsoon (Oct-Nov)', [10, 11],          61),
]

# --- DAYPART DEFINITION ---
# Each entry: (code, label, start_hour, end_hour)
# Rules:
#   - Hours are integers 0–24
#   - Entries must be contiguous: end of one = start of next
#   - First entry starts at 0, last entry ends at 24
#   - No gaps, no overlaps
# Scheme: 5dp_D_6_17 — 5 dp (solar 6-17, tight)
# Selected via sensitivity_timeslice_sweep.py ranking (April 2026).
# MUST match DAYPART_DEF in build_ostram_timeslices.py.
DAYPART_DEF = [
    ('D1', 'Night',         0,  6),
    ('D2', 'Solar day',     6, 17),
    ('D3', 'Evening peak', 17, 20),
    ('D4', 'Late evening', 20, 22),
    ('D5', 'Late night',   22, 24),
]

# ╔══════════════════════════════════════════════════════════════════════╗
# ║  DERIVED CONSTANTS — auto-generated from SEASON_DEF + DAYPART_DEF ║
# ║  DO NOT EDIT below this line.                                      ║
# ╚══════════════════════════════════════════════════════════════════════╝

def _validate_and_build():
    """Validate config and build all derived dicts. Called once at import."""
    # --- Seasons ---
    season_codes  = [s[0] for s in SEASON_DEF]
    season_names  = {s[0]: s[1] for s in SEASON_DEF}
    season_days   = {s[0]: s[3] for s in SEASON_DEF}
    month_to_season = {}
    for code, label, months, days in SEASON_DEF:
        for m in months:
            if m in month_to_season:
                raise ValueError(f"Month {m} assigned to both "
                                 f"{month_to_season[m]} and {code}")
            month_to_season[m] = code
    missing = [m for m in range(1, 13) if m not in month_to_season]
    if missing:
        raise ValueError(f"Months not assigned to any season: {missing}")
    if sum(season_days.values()) != 365:
        raise ValueError(f"Season days sum to {sum(season_days.values())}, not 365")

    # --- Dayparts ---
    daypart_codes = [d[0] for d in DAYPART_DEF]
    daypart_names = {d[0]: f"{d[1]} ({d[2]:02d}-{d[3]:02d})" for d in DAYPART_DEF}
    daypart_hours = {d[0]: d[3] - d[2] for d in DAYPART_DEF}
    daypart_ranges = {d[0]: (d[2], d[3]) for d in DAYPART_DEF}

    # Contiguity checks
    if DAYPART_DEF[0][2] != 0:
        raise ValueError(f"First daypart must start at hour 0, got {DAYPART_DEF[0][2]}")
    if DAYPART_DEF[-1][3] != 24:
        raise ValueError(f"Last daypart must end at hour 24, got {DAYPART_DEF[-1][3]}")
    for i in range(1, len(DAYPART_DEF)):
        if DAYPART_DEF[i][2] != DAYPART_DEF[i-1][3]:
            raise ValueError(f"Gap/overlap between {DAYPART_DEF[i-1][0]} "
                             f"(ends {DAYPART_DEF[i-1][3]}) and "
                             f"{DAYPART_DEF[i][0]} (starts {DAYPART_DEF[i][2]})")
    if sum(daypart_hours.values()) != 24:
        raise ValueError(f"Daypart hours sum to {sum(daypart_hours.values())}, not 24")

    # --- Timeslices ---
    timeslices = [s + d for s in season_codes for d in daypart_codes]

    # --- YearSplit ---
    yearsplit = {}
    for s_code, s_days in season_days.items():
        for d_code, d_hours in daypart_hours.items():
            yearsplit[s_code + d_code] = round((s_days * d_hours) / 8760, 6)

    # --- Hour-to-daypart lookup (hour 0–23 → daypart code) ---
    hour_daypart_lut = {}
    for d_code, d_label, d_start, d_end in DAYPART_DEF:
        for h in range(d_start, d_end):
            hour_daypart_lut[h] = d_code

    return {
        'season_codes': season_codes,
        'season_names': season_names,
        'season_days': season_days,
        'month_to_season': month_to_season,
        'daypart_codes': daypart_codes,
        'daypart_names': daypart_names,
        'daypart_hours': daypart_hours,
        'daypart_ranges': daypart_ranges,
        'timeslices': timeslices,
        'yearsplit': yearsplit,
        'hour_daypart_lut': hour_daypart_lut,
    }


_CFG = _validate_and_build()

SEASONS_LIST    = _CFG['season_codes']
SEASON_NAMES    = _CFG['season_names']
SEASON_DAYS_365 = _CFG['season_days']
MONTH_TO_SEASON = _CFG['month_to_season']
DAYPARTS_LIST   = _CFG['daypart_codes']
DAYPART_NAMES   = _CFG['daypart_names']
DAYPART_HOURS   = _CFG['daypart_hours']
DAYPART_RANGES  = _CFG['daypart_ranges']
TIMESLICES      = _CFG['timeslices']
YEARSPLIT_365   = _CFG['yearsplit']
_HOUR_LUT       = _CFG['hour_daypart_lut']

N_SEASONS       = len(SEASONS_LIST)
N_DAYPARTS      = len(DAYPARTS_LIST)
N_TIMESLICES    = len(TIMESLICES)


def hour_to_daypart(hour):
    """Map local hour (0–23) to daypart code using configured boundaries."""
    return _HOUR_LUT[int(hour) % 24]


def get_config_snapshot():
    """Return a serializable snapshot of the timeslice config for export."""
    return {
        'seasons': [{'code': s[0], 'label': s[1], 'months': s[2], 'days': s[3]}
                    for s in SEASON_DEF],
        'dayparts': [{'code': d[0], 'label': d[1], 'start_hour': d[2], 'end_hour': d[3]}
                     for d in DAYPART_DEF],
        'n_timeslices': N_TIMESLICES,
        'timeslices': TIMESLICES,
        'yearsplit': YEARSPLIT_365,
    }


# ============================================================================
# FUNCTIONS — file parsing
# ============================================================================

def parse_ninja_header(filepath):
    """Read the first 3 comment lines of a Renewables.ninja CSV."""
    meta = {}
    with open(filepath, 'r', encoding='utf-8-sig') as f:
        lines = [f.readline().strip() for _ in range(3)]

    desc = lines[0].lstrip('# ').strip()
    meta['header_description'] = desc

    json_line = lines[2]
    json_start = json_line.index('{')
    json_str   = json_line[json_start:]
    payload = json.loads(json_str)

    if 'units' in payload:
        for k, v in payload['units'].items():
            meta['unit_' + k] = v
    if 'params' in payload:
        for k, v in payload['params'].items():
            meta['param_' + k] = v

    return meta


def parse_filename(filepath):
    """Extract resource type, region code, and year from filename."""
    basename = os.path.splitext(os.path.basename(filepath))[0]
    parts = basename.split('_')
    info = {}
    if len(parts) >= 3:
        info['resource']  = parts[0]
        info['region']    = parts[1]
        info['year_file'] = parts[2]
    else:
        info['resource']  = parts[0] if len(parts) > 0 else 'unknown'
        info['region']    = parts[1] if len(parts) > 1 else 'unknown'
        info['year_file'] = 'unknown'
    return info


# ============================================================================
# FUNCTIONS — UTC offset and timeslice mapping
# ============================================================================

def get_utc_offset(region, param_lon):
    """Return UTC offset in hours. Dict overrides, else auto from longitude."""
    if region in UTC_OFFSETS:
        return UTC_OFFSETS[region]
    try:
        lon = float(param_lon)
        return round(lon / 15)
    except (ValueError, TypeError):
        print(f"  WARNING: cannot determine UTC offset for region={region}, "
              f"lon={param_lon}. Using UTC+0.")
        return 0


def assign_timeslice(df, utc_offset):
    """
    Shift timestamps UTC → local, then assign season, daypart, timeslice,
    and model_year. Uses the configured DAYPART_DEF boundaries.
    """
    df['utc_offset']  = utc_offset
    df['time_local']  = df['time'] + pd.Timedelta(hours=utc_offset)
    df['year_local']  = df['time_local'].dt.year
    df['month_local'] = df['time_local'].dt.month
    df['hour_local']  = df['time_local'].dt.hour
    df['doy_local']   = df['time_local'].dt.dayofyear
    df['season']      = df['month_local'].map(MONTH_TO_SEASON)
    df['daypart']     = df['hour_local'].apply(hour_to_daypart)
    df['timeslice']   = df['season'] + df['daypart']
    df['model_year']  = np.where(df['month_local'] == 12,
                                  df['year_local'] + 1,
                                  df['year_local'])
    df['season_name']  = df['season'].map(SEASON_NAMES)
    df['daypart_name'] = df['daypart'].map(DAYPART_NAMES)
    return df


# ============================================================================
# FUNCTIONS — read and collect
# ============================================================================

def read_ninja_csv(filepath):
    """Read a single Ninja CSV, attach metadata + timeslice columns."""
    meta_header   = parse_ninja_header(filepath)
    meta_filename = parse_filename(filepath)
    meta = {**meta_filename, **meta_header}

    df = pd.read_csv(filepath, comment='#', parse_dates=['time'])
    df.rename(columns={'electricity': 'cf'}, inplace=True)

    for k, v in meta.items():
        df[k] = v

    df['year_utc']  = df['time'].dt.year
    df['month_utc'] = df['time'].dt.month
    df['hour_utc']  = df['time'].dt.hour

    utc_offset = get_utc_offset(meta.get('region', ''),
                                meta.get('param_lon', None))
    df = assign_timeslice(df, utc_offset)

    return df, meta


def collect_files(solar_dir, wind_dir):
    """Glob both directories for *.csv and return a unified file list."""
    files = []
    for d in [solar_dir, wind_dir]:
        if os.path.isdir(d):
            found = sorted(glob.glob(os.path.join(d, '*.csv')))
            files.extend(found)
            print(f"  {d}: {len(found)} files found")
        else:
            print(f"  WARNING — directory not found: {d}")
    return files


# ============================================================================
# FUNCTIONS — timeslice aggregation
# ============================================================================

def aggregate_to_timeslices(df_hourly):
    """
    Compute mean CF per (resource, region, model_year, timeslice).
    Only model_years with >= 8000 hours are kept.
    """
    gk = ['resource', 'region', 'model_year', 'timeslice',
          'season', 'daypart', 'season_name', 'daypart_name']

    agg = df_hourly.groupby(gk).agg(
        cf_mean  = ('cf', 'mean'),
        cf_std   = ('cf', 'std'),
        cf_max   = ('cf', 'max'),
        cf_min   = ('cf', 'min'),
        n_hours  = ('cf', 'count'),
    ).reset_index()

    first_rec = (df_hourly.groupby(['resource', 'region'])
                 [['param_lat', 'param_lon', 'utc_offset']].first().reset_index())
    agg = agg.merge(first_rec, on=['resource', 'region'], how='left')

    agg['yearsplit'] = agg['timeslice'].map(YEARSPLIT_365)

    for c in ['cf_mean', 'cf_std', 'cf_max', 'cf_min']:
        agg[c] = agg[c].round(6)

    hours_per_my = (agg.groupby(['resource', 'region', 'model_year'])['n_hours']
                    .sum().reset_index().rename(columns={'n_hours': 'total_hours'}))
    agg = agg.merge(hours_per_my, on=['resource', 'region', 'model_year'])

    incomplete = agg[agg['total_hours'] < 8000][
        ['resource', 'region', 'model_year', 'total_hours']
    ].drop_duplicates()
    if len(incomplete) > 0:
        print("\n  Incomplete model_years (< 8000 hours, EXCLUDED):")
        for _, r in incomplete.iterrows():
            print(f"    {r['resource']:6s}  {r['region']:8s}  "
                  f"model_year={int(r['model_year'])}  "
                  f"hours={int(r['total_hours'])}")

    agg = agg[agg['total_hours'] >= 8000].copy()
    agg.drop(columns='total_hours', inplace=True)

    # Sort by canonical timeslice order
    ts_order = {t: i for i, t in enumerate(TIMESLICES)}
    agg['_ts_sort'] = agg['timeslice'].map(ts_order)
    agg.sort_values(['resource', 'region', 'model_year', '_ts_sort'], inplace=True)
    agg.drop(columns='_ts_sort', inplace=True)
    agg.reset_index(drop=True, inplace=True)

    return agg


# ============================================================================
# MAIN
# ============================================================================

if __name__ == '__main__':

    _T_WALL_START = time.perf_counter()
    _T_STEPS = {}

    print("=" * 70)
    print("Renewables.ninja CSV consolidation + OSeMOSYS timeslice mapping")
    print(f"Started: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 70)
    print(f"\nTimeslice structure: {N_SEASONS} seasons × {N_DAYPARTS} dayparts "
          f"= {N_TIMESLICES} timeslices")
    print(f"  Seasons : {SEASON_NAMES}")
    print(f"  Dayparts: {DAYPART_NAMES}")
    print(f"  Dec of calendar year Y → S1 of model_year Y+1")
    print(f"\nUTC handling:")
    print(f"  Auto-detect from longitude: offset = round(lon / 15)")
    if UTC_OFFSETS:
        print(f"  Manual overrides: {UTC_OFFSETS}")
    else:
        print(f"  No manual overrides set")

    # 1. Discover files
    _t0 = time.perf_counter()
    print("\nScanning directories...")
    all_files = collect_files(SOLAR_DIR, WIND_DIR)
    print(f"\nTotal files to process: {len(all_files)}")
    _T_STEPS['File discovery'] = time.perf_counter() - _t0

    if len(all_files) == 0:
        raise FileNotFoundError(
            "No CSV files found. Check SOLAR_DIR and WIND_DIR paths.")

    # 2. Read and stack
    _t0 = time.perf_counter()
    frames = []
    catalog_rows = []

    for i, fpath in enumerate(all_files):
        try:
            df_i, meta_i = read_ninja_csv(fpath)
            frames.append(df_i)

            cat = {**meta_i}
            cat['n_rows']      = len(df_i)
            cat['cf_mean']     = round(df_i['cf'].mean(), 4)
            cat['cf_max']      = round(df_i['cf'].max(), 4)
            cat['utc_offset']  = df_i['utc_offset'].iloc[0]
            cat['filepath']    = fpath
            catalog_rows.append(cat)

            if (i + 1) % 50 == 0 or (i + 1) == len(all_files):
                _elapsed = time.perf_counter() - _t0
                print(f"  processed {i + 1} / {len(all_files)}  "
                      f"({_elapsed:.1f}s elapsed)")

        except Exception as e:
            print(f"  ERROR reading {fpath}: {e}")
    _T_STEPS['Read CSVs'] = time.perf_counter() - _t0

    # 3. Concatenate hourly
    _t0 = time.perf_counter()
    print("\nConcatenating hourly data...")
    df_all = pd.concat(frames, ignore_index=True)
    df_catalog = pd.DataFrame(catalog_rows)

    lead_cols = ['time', 'time_local', 'cf', 'resource', 'region',
                 'year_local', 'model_year', 'month_local', 'hour_local',
                 'doy_local', 'year_file', 'utc_offset',
                 'timeslice', 'season', 'daypart', 'season_name', 'daypart_name',
                 'year_utc', 'month_utc', 'hour_utc']
    other_cols = [c for c in df_all.columns if c not in lead_cols]
    df_all = df_all[lead_cols + other_cols]
    _T_STEPS['Concatenate'] = time.perf_counter() - _t0

    # 4. Timeslice aggregation
    _t0 = time.perf_counter()
    print("\nAggregating to OSeMOSYS timeslices...")
    df_ts = aggregate_to_timeslices(df_all)
    _T_STEPS['Aggregation'] = time.perf_counter() - _t0

    # 5. Export
    _t0 = time.perf_counter()
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    out_hourly  = os.path.join(OUTPUT_DIR, 'compiled_reninja_hourly.csv')
    out_ts      = os.path.join(OUTPUT_DIR, 'compiled_reninja_ts.csv')
    out_catalog = os.path.join(OUTPUT_DIR, 'catalog_reninja.csv')
    out_config  = os.path.join(OUTPUT_DIR, 'timeslice_config.json')

    df_all.to_csv(out_hourly, index=False)
    df_ts.to_csv(out_ts, index=False)
    df_catalog.to_csv(out_catalog, index=False)

    with open(out_config, 'w') as f:
        json.dump(get_config_snapshot(), f, indent=2)
    _T_STEPS['Export CSVs'] = time.perf_counter() - _t0

    print(f"\nFiles written:")
    print(f"  Hourly table   : {out_hourly}")
    print(f"    shape        : {df_all.shape[0]:,} rows × {df_all.shape[1]} columns")
    print(f"  Timeslice table: {out_ts}")
    print(f"    shape        : {df_ts.shape[0]:,} rows × {df_ts.shape[1]} columns")
    print(f"  Catalog        : {out_catalog}")
    print(f"    files logged : {len(df_catalog)}")
    print(f"  Config snapshot: {out_config}")

    # 6. Summary
    print("\n" + "=" * 70)
    print("DATA SUMMARY — TIMESLICE AGGREGATED")
    print("=" * 70)
    print(f"\nResources   : {sorted(df_ts['resource'].unique())}")
    print(f"Regions     : {sorted(df_ts['region'].unique())}")
    print(f"Model years : {sorted(df_ts['model_year'].unique())}")

    offsets = df_catalog[['region', 'utc_offset']].drop_duplicates()
    print(f"\nUTC offsets applied:")
    for _, r in offsets.iterrows():
        print(f"  {r['region']:8s}  UTC{r['utc_offset']:+.1f}")

    if len(df_ts) > 0:
        print(f"\nSample: mean CF by timeslice (first resource × region × model_year):")
        first = df_ts.iloc[0]
        mask = ((df_ts['resource'] == first['resource']) &
                (df_ts['region'] == first['region']) &
                (df_ts['model_year'] == first['model_year']))
        sample = df_ts[mask][['timeslice', 'season_name', 'daypart_name',
                              'cf_mean', 'n_hours', 'yearsplit']]
        print(f"  {first['resource']} / {first['region']} / "
              f"model_year {int(first['model_year'])} / UTC{first['utc_offset']:+.1f}")
        print(sample.to_string(index=False))

    print(f"\n\nYearSplit reference ({sum(SEASON_DAYS_365.values())}-day year, "
          f"{N_SEASONS}S × {N_DAYPARTS}D = {N_TIMESLICES} ts):")
    ys_sum = 0
    for s in SEASONS_LIST:
        for d in DAYPARTS_LIST:
            ts = s + d
            ys = YEARSPLIT_365[ts]
            ys_sum += ys
            print(f"  {ts}  {SEASON_NAMES[s]:25s}  {DAYPART_NAMES[d]:22s}  "
                  f"YearSplit = {ys:.6f}")
    print(f"  {'TOTAL':51s}  {ys_sum:.6f}")

    # 7. Timing report
    _T_WALL_TOTAL = time.perf_counter() - _T_WALL_START

    print(f"\n{'='*70}")
    print(f"TIMING REPORT  ({N_SEASONS}S x {N_DAYPARTS}D = {N_TIMESLICES} ts, "
          f"{len(all_files)} files)")
    print(f"{'='*70}")
    for step, secs in _T_STEPS.items():
        if secs >= 60:
            print(f"  {step:<25s}  {secs/60:>6.1f} min  ({secs:>8.1f} s)")
        else:
            print(f"  {step:<25s}  {secs:>6.1f} s")
    print(f"  {'-'*45}")
    print(f"  {'TOTAL WALL TIME':<25s}  {_T_WALL_TOTAL/60:>6.1f} min  ({_T_WALL_TOTAL:>8.1f} s)")
    print()
    print(f"Config: {N_DAYPARTS} dayparts x {N_SEASONS} seasons = {N_TIMESLICES} timeslices")
    print(f"Finished: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 70)
