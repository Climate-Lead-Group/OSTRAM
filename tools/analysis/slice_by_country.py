"""
slice_by_country.py
───────────────────
Reads the OSTRAM combined inputs/outputs CSV, parses country and region codes
from the TECHNOLOGY column, and writes a filtered slice to disk.

Usage
-----
Set SELECTED_COUNTRY (and optionally SELECTED_REGION) in the config block
below, then run.  The script adds two columns to every row:

    COUNTRY  – 3-letter code(s) found in the tech name  (e.g. "BGD")
    REGION   – 5-char region code(s)                     (e.g. "BGDXX", "INDEA")

For cross-border TRN techs that touch two countries (e.g. TRNBGDXXINDEA),
both are listed pipe-separated: COUNTRY = "BGD|IND", REGION = "BGDXX|INDEA".
Filtering keeps any row where the selected country/region appears in either
position.

Naming anatomy
--------------
PWR/DSP/RNW/PWRSDS/PWRLDS : PREFIX(6) + REGION(5)    → PWRBIOBGDXX
MIN                        : MIN(3) + FUEL(3) + CC(3)  → MINCOABGD
TRN cross-border           : TRN(3) + REG(5) + REG(5)  → TRNBGDXXINDEA
TRN domestic (NLI/RPO)     : TRN(3) + TYPE(3) + REG(5) → TRNNLIBGDXX
"""

import os
import pandas as pd

# ═══════════════════════════════════════════════════════════════════════
# CONFIG  — edit these, then run
# ═══════════════════════════════════════════════════════════════════════
INPUT_CSV = r"OSTRAM_Combined_Inputs_Outputs.csv"   # path to the combined CSV
SELECTED_COUNTRY = "IND"       # BGD | BTN | IND | LKA | MDV | NPL | INT
SELECTED_REGION  = None        # None  → all sub-regions for that country
                               # or e.g. "INDEA", "INDNE", "INDNO", "INDSO", "INDWE"
SELECTED_SCENARIO = None       # None  → keep all scenarios
                               # or e.g. "A_Calibrated_BAU", "B_Optimised_VRE", "BAU"
OUTPUT_DIR = "."               # where to write the filtered CSV
# ═══════════════════════════════════════════════════════════════════════


# ── reference tables ──────────────────────────────────────────────────
REGIONS = [
    "BGDXX", "BTNXX",
    "INDEA", "INDNE", "INDNO", "INDSO", "INDWE",
    "LKAXX", "MDVXX", "NPLXX",
]

COUNTRIES = ["BGD", "BTN", "IND", "LKA", "MDV", "NPL", "INT"]

COUNTRY_LABELS = {
    "BGD": "Bangladesh",
    "BTN": "Bhutan",
    "IND": "India",
    "LKA": "Sri Lanka",
    "MDV": "Maldives",
    "NPL": "Nepal",
    "INT": "International",
}


# ── parsing logic ─────────────────────────────────────────────────────
def parse_regions(tech: str) -> list[str]:
    """Return all 5-char region codes found in a technology name."""
    return [r for r in REGIONS if r in tech]


def parse_countries(tech: str) -> list[str]:
    """Return all 3-char country codes found in a technology name.

    Uses the 5-char regions first (to avoid false positives from substring
    overlaps), then falls back to bare 3-char codes for MIN*/INT techs.
    """
    found_regions = parse_regions(tech)
    found_countries = list(dict.fromkeys(r[:3] for r in found_regions))  # order-preserving unique

    # MIN techs carry a bare 3-char country code with no sub-region suffix
    for cc in COUNTRIES:
        if cc not in found_countries and cc in tech:
            found_countries.append(cc)

    return found_countries


# ── main ──────────────────────────────────────────────────────────────
def main():
    print(f"Reading {INPUT_CSV} …")
    df = pd.read_csv(INPUT_CSV, low_memory=False)
    print(f"  {len(df):,} rows × {len(df.columns)} columns loaded.")

    # strip Windows line-ending artefacts from the Scenario column
    if "Scenario" in df.columns:
        df["Scenario"] = df["Scenario"].astype(str).str.strip()

    # parse country / region from TECHNOLOGY
    print("Parsing country & region codes …")
    df["COUNTRY"] = df["TECHNOLOGY"].fillna("").apply(lambda t: "|".join(parse_countries(t)) or "UNKNOWN")
    df["REGION"]  = df["TECHNOLOGY"].fillna("").apply(lambda t: "|".join(parse_regions(t))   or "")

    # ── quick summary ─────────────────────────────────────────────────
    all_techs = df["TECHNOLOGY"].nunique()
    print(f"\n  Unique technologies: {all_techs}")
    print(f"  Scenarios:           {sorted(df['Scenario'].unique()) if 'Scenario' in df.columns else 'n/a'}")
    print("\n  Rows by country (a tech touching two countries counts in both):")
    for cc in COUNTRIES:
        n = df["COUNTRY"].str.contains(cc, na=False).sum()
        label = COUNTRY_LABELS.get(cc, cc)
        print(f"    {cc}  {label:<15s}  {n:>9,} rows")

    # ── filter ────────────────────────────────────────────────────────
    mask = pd.Series(True, index=df.index)

    if SELECTED_COUNTRY:
        mask &= df["COUNTRY"].str.contains(SELECTED_COUNTRY, na=False)
        print(f"\n→ Filtered to COUNTRY containing '{SELECTED_COUNTRY}'  ({mask.sum():,} rows)")

    if SELECTED_REGION:
        mask &= df["REGION"].str.contains(SELECTED_REGION, na=False)
        print(f"→ Filtered to REGION  containing '{SELECTED_REGION}'   ({mask.sum():,} rows)")

    if SELECTED_SCENARIO:
        mask &= df["Scenario"] == SELECTED_SCENARIO
        print(f"→ Filtered to Scenario == '{SELECTED_SCENARIO}'         ({mask.sum():,} rows)")

    out = df.loc[mask].copy()

    # ── write ─────────────────────────────────────────────────────────
    parts = ["OSTRAM_slice", SELECTED_COUNTRY or "ALL"]
    if SELECTED_REGION:
        parts.append(SELECTED_REGION)
    if SELECTED_SCENARIO:
        parts.append(SELECTED_SCENARIO)
    out_name = "_".join(parts) + ".csv"
    out_path = os.path.join(OUTPUT_DIR, out_name)

    out.to_csv(out_path, index=False)
    print(f"\n✓ Wrote {len(out):,} rows → {out_path}")
    print(f"  Unique techs in slice: {out['TECHNOLOGY'].nunique()}")
    if SELECTED_COUNTRY == "IND" and not SELECTED_REGION:
        print("  Sub-regions in slice:", sorted(out["REGION"].str.split("|").explode().unique()))


if __name__ == "__main__":
    main()
