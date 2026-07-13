# -*- coding: utf-8 -*-
"""
WS-3 Phase 2 — INTERNAL transmission residual capacities via the peak x 1.2 method.

PURE CALCULATION. Emits a desk-check CSV. NO model mutation.

Per node: size existing internal transmission at (peak demand x 1.2), split into a
renewable-carrying share (RNWTRN) and non-renewable-carrying share (PWRTRN) using the
generation available at the node's peak timeslice. NLI/RPO families stay 0.
Interconnector residuals are out of scope here (handled from the physical project list).

Reference year 2023 (base year). Reads the post-B1 otoole params from the workcopy.
"""
from __future__ import annotations
import os, re
from datetime import datetime
import pandas as pd

# --------------------------------------------------------------------- config / constants
BASE = r"C:\Users\luisfernando\Desktop\OSeMOSYS\OSTRAM_ws3_workcopy\t1_confection\A2_Output_Params\A_Calibrated_BAU"
OUTDIR = r"C:\Users\luisfernando\Desktop\OSeMOSYS\OSTRAM_ws3_workcopy\ws3_transmission_audit\outputs"
YEAR = 2023
MARGIN = 1.2                    # transmission headroom on peak (Javier convention; NOT the 1.15 gen ReserveMargin)
PJ_TO_MWH = 277777.8
HOURS_PER_YEAR = 8760
FLAT_PLACEHOLDER = 5.0          # current model value for every transmission family/node

RE_CODES = {"BIO", "HYD", "CSP", "GEO", "SPV", "WAS", "WON", "WOF"}   # Config renewable_fuels
VARIABLE_RE = {"SPV", "WON", "WOF", "CSP", "HYD"}                     # use per-timeslice CF
ALWAYS_RE = {"BIO", "GEO", "WAS"}                                    # firm RE -> CF = 1.0
STORAGE = {"SDS", "LDS", "HPS"}                                      # shifting devices -> excluded from generation

NODES = ["BGDXX", "BTNXX", "INDEA", "INDNE", "INDNO", "INDSO", "INDWE", "LKAXX", "MDVXX", "NPLXX"]

# --------------------------------------------------------------------- load helpers
def load(name):
    df = pd.read_csv(os.path.join(BASE, name + ".csv"))
    df.columns = [c.strip() for c in df.columns]
    return df

def C(df, *names):
    up = {c.upper(): c for c in df.columns}
    for n in names:
        if n.upper() in up:
            return up[n.upper()]
    raise KeyError(f"none of {names} in {list(df.columns)}")

def y23(df):
    return df[df[C(df, "YEAR")].astype(float) == YEAR]

def energy_of(tech):   # PWR + energy + node(last5)
    return tech[3:-5]

def node_of(tech):
    return tech[-5:]

# --------------------------------------------------------------------- load inputs
sad = y23(load("SpecifiedAnnualDemand"))
sdp = y23(load("SpecifiedDemandProfile"))
ys = y23(load("YearSplit"))
rc = y23(load("ResidualCapacity"))
cf = y23(load("CapacityFactor"))

# YearSplit: {ts: frac}
ys_ts, ys_v = C(ys, "TIMESLICE"), C(ys, "VALUE")
yearsplit = dict(zip(ys[ys_ts], ys[ys_v]))

# Demand: {fuel: PJ}
sad_f, sad_v = C(sad, "FUEL"), C(sad, "VALUE")
demand = dict(zip(sad[sad_f], sad[sad_v]))

# Profile: {(fuel, ts): frac}
sdp_f, sdp_ts, sdp_v = C(sdp, "FUEL"), C(sdp, "TIMESLICE"), C(sdp, "VALUE")
profile = {(r[sdp_f], r[sdp_ts]): r[sdp_v] for _, r in sdp.iterrows()}

# ResidualCapacity PWR: {tech: GW}
rc_t, rc_v = C(rc, "TECHNOLOGY"), C(rc, "VALUE")
rc_pwr = rc[rc[rc_t].astype(str).str.startswith("PWR")]
rescap = {r[rc_t]: r[rc_v] for _, r in rc_pwr.iterrows()}

# CapacityFactor: {(tech, ts): cf}
cf_t, cf_ts, cf_v = C(cf, "TECHNOLOGY"), C(cf, "TIMESLICE"), C(cf, "VALUE")
cfmap = {(r[cf_t], r[cf_ts]): r[cf_v] for _, r in cf.iterrows()}
cf_by_tech = {}
for (t, ts), v in cfmap.items():
    cf_by_tech.setdefault(t, []).append(v)

# --------------------------------------------------------------------- energy-code audit
print("=" * 74)
print("ENERGY-CODE AUDIT (PWR techs with 2023 ResidualCapacity)")
codes = sorted({energy_of(t) for t in rescap})
def classify(e):
    if e in STORAGE: return "STORAGE (excluded)"
    if e in ALWAYS_RE: return "RE firm (CF=1.0)"
    if e in VARIABLE_RE: return "RE variable (CF@peak)"
    if e in RE_CODES: return "RE (other)"
    return "non-RE"
for e in codes:
    print(f"   {e:5} -> {classify(e)}")
unknown = [e for e in codes if e not in RE_CODES and e not in STORAGE]
print(f"   (non-RE codes treated as residual, not summed: {unknown})")

# --------------------------------------------------------------------- per-node calc
def cf_at(tech, ts):
    if (tech, ts) in cfmap:
        return cfmap[(tech, ts)]
    vals = cf_by_tech.get(tech)
    return (sum(vals) / len(vals)) if vals else 1.0   # annual-avg fallback, else 1.0

rows = []
for n in NODES:
    fuel = f"ELC{n}03"
    dem = demand.get(fuel)
    if dem is None:
        rows.append({"node": n, "status_flags": "NO_DEMAND_FUEL"}); continue
    # 1. peak timeslice / peak power
    powers = {}
    for ts, frac in yearsplit.items():
        pf = profile.get((fuel, ts), 0.0)
        powers[ts] = dem * pf * PJ_TO_MWH / (frac * HOURS_PER_YEAR) / 1000.0  # GW
    ts_star = max(powers, key=powers.get)
    peak_GW = powers[ts_star]
    # 2. RE available at peak
    re_avail = 0.0
    for t, cap in rescap.items():
        if node_of(t) != n:
            continue
        e = energy_of(t)
        if e in STORAGE or e not in RE_CODES:
            continue
        cfac = cf_at(t, ts_star) if e in VARIABLE_RE else (1.0 if e in ALWAYS_RE else cf_at(t, ts_star))
        re_avail += cap * cfac
    # 3/4. split with 1.2 margin
    rnwtrn = re_avail * MARGIN
    pwrtrn = peak_GW * MARGIN - re_avail * MARGIN
    flags = []
    if pwrtrn < 0:
        flags.append("RE_SATURATES(PWRTRN floored 0; RNWTRN capped)")
        rnwtrn = peak_GW * MARGIN
        pwrtrn = 0.0
    if peak_GW < 1.5:
        flags.append("SMALL_GRID")
    if re_avail == 0:
        flags.append("NO_RE_RESIDUAL")
    identity_ok = abs((rnwtrn + pwrtrn) - peak_GW * MARGIN) < 1e-6
    rows.append({
        "node": n, "peak_GW": round(peak_GW, 3), "peak_timeslice": ts_star,
        "RE_avail_GW": round(re_avail, 3), "re_share": round(re_avail / peak_GW, 3) if peak_GW else None,
        "RNWTRN_residual_GW": round(rnwtrn, 3), "PWRTRN_residual_GW": round(pwrtrn, 3),
        "identity_ok": identity_ok, "status_flags": "; ".join(flags) or "ok",
    })

df = pd.DataFrame(rows)
os.makedirs(OUTDIR, exist_ok=True)
stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
out = os.path.join(OUTDIR, f"internal_tx_residuals_{stamp}.csv")
df.to_csv(out, index=False)

# --------------------------------------------------------------------- report
print("\n" + "=" * 74)
print("PER-NODE INTERNAL TRANSMISSION RESIDUALS (peak x 1.2)")
print("=" * 74)
print(df.to_string(index=False))
print(f"\nIdentity RNWTRN+PWRTRN == peak x 1.2 holds for ALL nodes: {bool(df['identity_ok'].all())}")

print("\n" + "=" * 74)
print("BEFORE (flat placeholder) vs AFTER (computed) — GW")
print("=" * 74)
print(f"{'node':7} {'RNWTRN: flat':>13} {'computed':>10}   {'PWRTRN: flat':>13} {'computed':>10}")
for _, r in df.iterrows():
    if "peak_GW" not in r or pd.isna(r.get("peak_GW")):
        continue
    print(f"{r['node']:7} {FLAT_PLACEHOLDER:>13.1f} {r['RNWTRN_residual_GW']:>10.2f}   "
          f"{FLAT_PLACEHOLDER:>13.1f} {r['PWRTRN_residual_GW']:>10.2f}")
print("\nNLI/RPO families (RNWNLI, RNWRPO, TRNNLI, TRNRPO): ResidualCapacity = 0 (build mechanisms).")
print(f"\nCSV -> {out}")
