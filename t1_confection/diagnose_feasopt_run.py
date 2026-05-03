#!/usr/bin/env python3
"""
diagnose_feasopt_run.py

Parse a CPLEX FEASOPT .sol file + the BAU_0_Output.csv produced by B2.py,
and produce a structured diagnostic report:

  A. FEASOPT relaxation summary
     - total relaxation, # constraints relaxed
     - by constraint type (TCC1_TotalAnnualMaxCapacityConstraint, etc.)
     - by region / tech / year
     - top-N individual relaxations (largest |slack|)

  B. Build mix snapshot (NewCapacity by region × year-band)
     - early (2023-2030, 0.5% lid band)
     - mid   (2031-2040, 10% band)
     - late  (2041-2050, 50% band)

  C. Lid utilization (optional, requires --changes-json from the lid run)
     - histogram of NewCapacity / lid_value
     - top-N most binding cells (utilization closest to 1)
     - top-N "lid-busted" cells (NewCapacity > lid_value, i.e. FEASOPT relaxed it)

The script is read-only on inputs and only writes its --report-json output (if
specified). Does not touch the LP, the .sol, the CSV, or any pipeline artifact.

Usage:
  python diagnose_feasopt_run.py \
      --feasopt-sol Executables/BAU_0/Pre_processed_BAU_0_NoStorage_output.feasopt.sol \
      --output-csv  BAU_0_Output.csv \
      --changes-json A1_Outputs/A1_Outputs_BAU_PRE_LID_*/A-O_Parametrization.xlsx_CHANGES.json

The --changes-json is optional; without it section C is skipped.
"""

from __future__ import annotations
import argparse
import json
import re
import sys
from collections import defaultdict
from pathlib import Path
from xml.etree import ElementTree as ET


# ----- Constraint name parsing ------------------------------------------------

# OSeMOSYS constraint names look like:
#   TCC1_TotalAnnualMaxCapacityConstraint(GLOBAL,PWRWONNPLXX,2043)
#   CAa2_TotalAnnualCapacity(GLOBAL,DSPTRNBGDXX,2023)
#   EBa1_RateOfFuelProduction1(GLOBAL,ELCBGDXX01,S1,D1,B1,2023)
NAME_RE = re.compile(r"^([A-Za-z0-9_]+)\(([^)]*)\)$")
YEAR_RE = re.compile(r"^(20[2-5][0-9])$")


def parse_constraint_name(name: str):
    """Return (type_prefix, region, tech, year, raw_args). Best-effort:
    region defaults to GLOBAL if first arg looks region-like, tech is the
    arg matching the PWR/ELC/etc pattern, year is any 4-digit 20xx token."""
    m = NAME_RE.match(name)
    if not m:
        return (name, None, None, None, [])
    ctype = m.group(1)
    args = [a.strip() for a in m.group(2).split(",") if a.strip()]
    region = None
    tech = None
    year = None
    for a in args:
        if YEAR_RE.match(a):
            year = int(a)
        elif a == "GLOBAL" or (a.isupper() and len(a) <= 6 and not any(c.isdigit() for c in a)):
            if region is None:
                region = a
        elif re.match(r"^(PWR|ELC|DSP|TRN|MIN|REF|ENG)", a):
            if tech is None:
                tech = a
    return (ctype, region, tech, year, args)


def extract_country_region(tech_or_region: str | None) -> str | None:
    """For a tech like PWRWONNPLXX, return NPLXX. For a fuel like ELCNPLXX01,
    return NPLXX. Otherwise None."""
    if not tech_or_region:
        return None
    # PWRXXXYYYZZ -> chars 6:11 in the 11-char form (e.g. PWRWONNPLXX -> NPLXX)
    if len(tech_or_region) >= 11 and tech_or_region[:3] in {"PWR", "DSP", "MIN", "REF"}:
        return tech_or_region[6:11]
    # ELCXXXNN -> chars 3:8 (e.g. ELCNPLXX01 -> NPLXX)
    if len(tech_or_region) >= 8 and tech_or_region[:3] == "ELC":
        return tech_or_region[3:8]
    return None


# ----- FEASOPT .sol parser ----------------------------------------------------

def parse_feasopt_sol(path: Path):
    """Stream-parse the XML and yield (name, slack_float) for every constraint
    with non-zero slack. Also returns header metadata."""
    header = {}
    constraints_iter = []  # list of (name, slack)
    # Use iterparse to handle large files
    context = ET.iterparse(str(path), events=("start", "end"))
    for event, elem in context:
        if event == "start" and elem.tag == "header":
            header = dict(elem.attrib)
        elif event == "end" and elem.tag == "constraint":
            name = elem.attrib.get("name", "")
            slack_str = elem.attrib.get("slack", "0")
            try:
                slack = float(slack_str)
            except ValueError:
                slack = 0.0
            if slack != 0.0:
                constraints_iter.append((name, slack))
            elem.clear()  # free memory
    return header, constraints_iter


# ----- BAU_0_Output.csv parser ------------------------------------------------

def load_build_mix(csv_path: Path):
    """Load NewCapacity rows from BAU_0_Output.csv. Returns a dict
    {(tech, year): newcap_float} aggregating across whatever indexing the
    CSV has."""
    import pandas as pd
    # Read only the columns we need
    cols_want = ["TECHNOLOGY", "YEAR", "NewCapacity"]
    # The CSV may have other columns we don't need
    df = pd.read_csv(csv_path, usecols=lambda c: c in cols_want, low_memory=False)
    # Filter to rows with a tech, year, and NewCapacity value
    df = df.dropna(subset=["TECHNOLOGY", "YEAR", "NewCapacity"])
    df = df[df["NewCapacity"] != 0]
    # Coerce types
    df["YEAR"] = df["YEAR"].astype(int)
    df["NewCapacity"] = df["NewCapacity"].astype(float)
    # Aggregate (tech, year) -- sum because OSeMOSYS may emit duplicates
    # across timeslices/modes for the same NewCapacity record
    out = (df.groupby(["TECHNOLOGY", "YEAR"])["NewCapacity"]
             .max()  # take max not sum: NewCapacity is a per-(tech,year) variable
             .to_dict())
    return out


def year_band(year: int) -> str:
    if year <= 2030: return "early (2023-2030, 0.5%)"
    if year <= 2040: return "mid (2031-2040, 10%)"
    return "late (2041-2050, 50%)"


# ----- Reports ---------------------------------------------------------------

def fmt_num(x):
    if x is None:
        return "—"
    if abs(x) >= 1000:
        return f"{x:,.0f}"
    if abs(x) >= 1:
        return f"{x:,.3f}"
    return f"{x:.5f}"


def report_a_feasopt(header, relaxations, top_n):
    print("=" * 78)
    print("A. FEASOPT RELAXATION SUMMARY")
    print("=" * 78)
    obj = float(header.get("objectiveValue", 0))
    status = header.get("solutionStatusString", "")
    print(f"  Status: {status}")
    print(f"  Total relaxation (objectiveValue): {fmt_num(obj)}")
    print(f"  Constraints with non-zero slack:   {len(relaxations)}")
    if not relaxations:
        print("  (No relaxed constraints — no infeasibility to diagnose.)")
        return

    # Parse all
    parsed = []
    for name, slack in relaxations:
        ctype, region, tech, year, _ = parse_constraint_name(name)
        country = extract_country_region(tech) or extract_country_region(region) or region
        parsed.append({
            "name": name, "slack": slack, "abs_slack": abs(slack),
            "ctype": ctype, "country": country, "tech": tech, "year": year,
        })

    # By constraint type
    print()
    print("  By constraint type (sorted by total |slack|):")
    by_type = defaultdict(lambda: {"count": 0, "abs_total": 0.0})
    for p in parsed:
        by_type[p["ctype"]]["count"] += 1
        by_type[p["ctype"]]["abs_total"] += p["abs_slack"]
    rows = sorted(by_type.items(), key=lambda kv: -kv[1]["abs_total"])
    print(f"    {'TYPE':<55} {'COUNT':>8} {'TOTAL_|SLACK|':>15}")
    for k, v in rows:
        print(f"    {k:<55} {v['count']:>8} {fmt_num(v['abs_total']):>15}")

    # By country
    print()
    print("  By country/region:")
    by_country = defaultdict(lambda: {"count": 0, "abs_total": 0.0})
    for p in parsed:
        c = p["country"] or "(none)"
        by_country[c]["count"] += 1
        by_country[c]["abs_total"] += p["abs_slack"]
    rows = sorted(by_country.items(), key=lambda kv: -kv[1]["abs_total"])
    print(f"    {'COUNTRY':<10} {'COUNT':>8} {'TOTAL_|SLACK|':>15}")
    for k, v in rows:
        print(f"    {k:<10} {v['count']:>8} {fmt_num(v['abs_total']):>15}")

    # By year
    print()
    print("  By year:")
    by_year = defaultdict(lambda: {"count": 0, "abs_total": 0.0})
    for p in parsed:
        y = p["year"] if p["year"] is not None else "(none)"
        by_year[y]["count"] += 1
        by_year[y]["abs_total"] += p["abs_slack"]
    rows = sorted(by_year.items(), key=lambda kv: (str(kv[0])))
    print(f"    {'YEAR':<8} {'COUNT':>8} {'TOTAL_|SLACK|':>15}")
    for k, v in rows:
        print(f"    {str(k):<8} {v['count']:>8} {fmt_num(v['abs_total']):>15}")

    # By tech (top 20 only)
    print()
    print("  By tech (top 20 by total |slack|):")
    by_tech = defaultdict(lambda: {"count": 0, "abs_total": 0.0})
    for p in parsed:
        t = p["tech"] or "(none)"
        by_tech[t]["count"] += 1
        by_tech[t]["abs_total"] += p["abs_slack"]
    rows = sorted(by_tech.items(), key=lambda kv: -kv[1]["abs_total"])[:20]
    print(f"    {'TECH':<14} {'COUNT':>8} {'TOTAL_|SLACK|':>15}")
    for k, v in rows:
        print(f"    {k:<14} {v['count']:>8} {fmt_num(v['abs_total']):>15}")

    # Top-N individual relaxations
    print()
    print(f"  Top {top_n} individual relaxed constraints (by |slack|):")
    parsed_sorted = sorted(parsed, key=lambda p: -p["abs_slack"])[:top_n]
    print(f"    {'#':>3}  {'NAME':<70} {'SLACK':>15}")
    for i, p in enumerate(parsed_sorted, 1):
        print(f"    {i:>3}  {p['name']:<70} {fmt_num(p['slack']):>15}")


def report_b_buildmix(builds, top_n):
    print()
    print("=" * 78)
    print("B. BUILD MIX SNAPSHOT (NewCapacity)")
    print("=" * 78)
    if not builds:
        print("  (No NewCapacity rows found in CSV.)")
        return

    # By country × band
    by_cb = defaultdict(lambda: defaultdict(float))
    by_cb_count = defaultdict(lambda: defaultdict(int))
    for (tech, year), nc in builds.items():
        country = extract_country_region(tech)
        if country is None: continue
        b = year_band(year)
        by_cb[country][b] += nc
        by_cb_count[country][b] += 1

    print()
    print("  Total NewCapacity by country × year-band (GW):")
    bands = ["early (2023-2030, 0.5%)", "mid (2031-2040, 10%)", "late (2041-2050, 50%)"]
    countries = sorted(by_cb.keys())
    print(f"    {'COUNTRY':<10} {bands[0]:>30} {bands[1]:>30} {bands[2]:>30}")
    for c in countries:
        row = [f"{fmt_num(by_cb[c].get(b, 0)):>30}" for b in bands]
        print(f"    {c:<10} {''.join(row)}")

    # Tech-share concentration in each (country, band)
    print()
    print("  Tech share concentration (top tech share per country × band):")
    # Aggregate (country, band, tech)
    by_cbt = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))
    for (tech, year), nc in builds.items():
        country = extract_country_region(tech)
        if country is None: continue
        tech_class = tech[3:6] if len(tech) >= 6 else tech  # PWRSPV -> SPV
        b = year_band(year)
        by_cbt[country][b][tech_class] += nc

    print(f"    {'COUNTRY':<10} {'BAND':<28} {'TOP TECH':<10} {'SHARE':>8} {'TOTAL':>12}")
    for c in countries:
        for b in bands:
            shares = by_cbt[c].get(b, {})
            total = sum(shares.values())
            if total == 0: continue
            top_tech, top_val = max(shares.items(), key=lambda kv: kv[1])
            share = top_val / total
            print(f"    {c:<10} {b:<28} {top_tech:<10} {share*100:>7.1f}% {fmt_num(total):>12}")

    # Top-N individual builds
    print()
    print(f"  Top {top_n} individual (tech, year) builds (largest NewCapacity, GW):")
    builds_sorted = sorted(builds.items(), key=lambda kv: -kv[1])[:top_n]
    print(f"    {'#':>3}  {'TECH':<14} {'YEAR':>6} {'NEWCAP':>14}")
    for i, ((t, y), nc) in enumerate(builds_sorted, 1):
        print(f"    {i:>3}  {t:<14} {y:>6} {fmt_num(nc):>14}")


def report_c_lid_utilization(builds, changes_json_path, top_n):
    print()
    print("=" * 78)
    print("C. LID UTILIZATION (NewCapacity / lid_value)")
    print("=" * 78)
    with open(changes_json_path) as f:
        changes = json.load(f)
    s = changes["sheets"][0]
    # Build (tech, year) -> lid_value from the 'changes' records
    lid = {}
    for c in s["changes"]:
        lid[(c["tech"], c["year"])] = c["new"]
    for p in s["preserved"]:
        lid[(p["tech"], p["year"])] = p["value"]
    print(f"  Lid grid loaded: {len(lid)} (tech, year) cells "
          f"covering {len(s['allowed_techs'])} allowed techs.")
    print(f"  Build cells available: {len(builds)} from CSV.")

    # Compute utilization where both lid and NewCapacity exist and lid > 0
    util = []
    for (tech, year), lid_val in lid.items():
        if lid_val is None or lid_val <= 0:
            continue
        nc = builds.get((tech, year), 0.0)
        if nc == 0:
            continue  # skip un-built cells (uninteresting for this view)
        util.append({
            "tech": tech, "year": year, "lid": lid_val, "newcap": nc,
            "ratio": nc / lid_val,
        })

    if not util:
        print("  No cells with both a lid and a non-zero NewCapacity.")
        return

    # Histogram
    buckets = [(0, 0.01), (0.01, 0.5), (0.5, 0.95), (0.95, 1.05),
               (1.05, 2.0), (2.0, float("inf"))]
    print()
    print(f"  Utilization histogram across {len(util)} built cells:")
    print(f"    {'BUCKET':<22} {'COUNT':>8} {'INTERPRETATION':<40}")
    interp = {
        (0, 0.01):    "near-zero (lid not active)",
        (0.01, 0.5):  "low slack (well below lid)",
        (0.5, 0.95):  "moderate (using <95% of lid)",
        (0.95, 1.05): "BINDING (at or near lid)",
        (1.05, 2.0):  "lid VIOLATED (FEASOPT relaxed)",
        (2.0, float("inf")): "lid badly violated",
    }
    for lo, hi in buckets:
        cnt = sum(1 for u in util if lo <= u["ratio"] < hi)
        label = f"[{lo:.2f}, {'inf' if hi == float('inf') else f'{hi:.2f}'})"
        print(f"    {label:<22} {cnt:>8} {interp[(lo, hi)]:<40}")

    # Top binding (closest to 1.0 from below)
    print()
    print(f"  Top {top_n} most-binding cells (ratio nearest 1.0, ratio <= 1.0):")
    binding = sorted([u for u in util if u["ratio"] <= 1.0],
                     key=lambda u: -u["ratio"])[:top_n]
    print(f"    {'#':>3}  {'TECH':<14} {'YEAR':>6} {'NEWCAP':>12} {'LID':>12} {'RATIO':>8}")
    for i, u in enumerate(binding, 1):
        print(f"    {i:>3}  {u['tech']:<14} {u['year']:>6} "
              f"{fmt_num(u['newcap']):>12} {fmt_num(u['lid']):>12} "
              f"{u['ratio']*100:>7.1f}%")

    # Top busted (ratio > 1)
    busted = sorted([u for u in util if u["ratio"] > 1.0],
                    key=lambda u: -u["ratio"])[:top_n]
    print()
    print(f"  Top {top_n} 'lid-busted' cells (NewCapacity > lid; FEASOPT relaxed):")
    if not busted:
        print(f"    (none — every built cell stayed within its lid)")
    else:
        print(f"    {'#':>3}  {'TECH':<14} {'YEAR':>6} {'NEWCAP':>12} {'LID':>12} {'RATIO':>8}")
        for i, u in enumerate(busted, 1):
            print(f"    {i:>3}  {u['tech']:<14} {u['year']:>6} "
                  f"{fmt_num(u['newcap']):>12} {fmt_num(u['lid']):>12} "
                  f"{u['ratio']*100:>7.1f}%")


def main():
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--feasopt-sol", required=True, type=Path,
                    help="Path to *.feasopt.sol XML file")
    ap.add_argument("--output-csv", required=True, type=Path,
                    help="Path to BAU_0_Output.csv")
    ap.add_argument("--changes-json", type=Path, default=None,
                    help="Optional: *_CHANGES.json from the lid run, enables lid utilization report")
    ap.add_argument("--top", type=int, default=20, help="Top-N depth for lists (default 20)")
    ap.add_argument("--report-json", type=Path, default=None,
                    help="Optional: write a structured JSON summary to this path")
    args = ap.parse_args()

    for p in [args.feasopt_sol, args.output_csv]:
        if not p.exists():
            sys.exit(f"Not found: {p}")
    if args.changes_json and not args.changes_json.exists():
        sys.exit(f"Not found: {args.changes_json}")

    print(f"Inputs:")
    print(f"  feasopt-sol:  {args.feasopt_sol}")
    print(f"  output-csv:   {args.output_csv}")
    print(f"  changes-json: {args.changes_json or '(skipped — section C disabled)'}")
    print()

    print("Parsing FEASOPT .sol ...")
    header, relaxations = parse_feasopt_sol(args.feasopt_sol)
    print(f"  -> {len(relaxations)} non-zero-slack constraints")

    print("Loading build mix from CSV ...")
    builds = load_build_mix(args.output_csv)
    print(f"  -> {len(builds)} (tech, year) cells with non-zero NewCapacity")

    report_a_feasopt(header, relaxations, args.top)
    report_b_buildmix(builds, args.top)
    if args.changes_json:
        report_c_lid_utilization(builds, args.changes_json, args.top)

    if args.report_json:
        # minimal structured dump
        out = {
            "header": header,
            "relaxations": [{"name": n, "slack": s} for n, s in relaxations],
            "builds": [{"tech": t, "year": y, "newcap": v}
                       for (t, y), v in builds.items()],
        }
        with open(args.report_json, "w") as f:
            json.dump(out, f, indent=1)
        print(f"\nStructured JSON written: {args.report_json}")

    print()
    print("Done.")


if __name__ == "__main__":
    main()
