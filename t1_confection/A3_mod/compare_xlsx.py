"""compare_xlsx.py — content-only diff between two .xlsx workbooks.

Usage:
    python compare_xlsx.py A.xlsx B.xlsx
Exit 0 if every sheet's cell content matches; non-zero otherwise.
"""
import sys
from pathlib import Path
import numpy as np
import pandas as pd


def compare(a: Path, b: Path, tol: float = 1e-9) -> dict:
    xa, xb = pd.ExcelFile(a), pd.ExcelFile(b)
    sheets = sorted(set(xa.sheet_names) | set(xb.sheet_names))
    report = {}
    for s in sheets:
        if s not in xa.sheet_names:
            report[s] = "missing in A"
            continue
        if s not in xb.sheet_names:
            report[s] = "missing in B"
            continue
        da = pd.read_excel(xa, s)
        db = pd.read_excel(xb, s)
        if da.shape != db.shape:
            report[s] = f"shape diff: A={da.shape} B={db.shape}"
            continue
        if list(da.columns) != list(db.columns):
            report[s] = f"column-name mismatch: A={list(da.columns)[:5]}... B={list(db.columns)[:5]}..."
            continue
        diffs = 0
        first_diff = None
        for col in da.columns:
            ca, cb = da[col], db[col]
            if pd.api.types.is_numeric_dtype(ca) and pd.api.types.is_numeric_dtype(cb):
                a_arr = ca.fillna(0).to_numpy(dtype=float)
                b_arr = cb.fillna(0).to_numpy(dtype=float)
                m = ~np.isclose(a_arr, b_arr, atol=tol, equal_nan=True)
            else:
                m = ca.fillna("__NA__").astype(str).to_numpy() != cb.fillna("__NA__").astype(str).to_numpy()
            n = int(m.sum())
            if n > 0 and first_diff is None:
                idx = int(np.argmax(m))
                first_diff = (col, idx, ca.iloc[idx], cb.iloc[idx])
            diffs += n
        if diffs == 0:
            report[s] = "OK"
        else:
            col, idx, va, vb = first_diff
            report[s] = f"{diffs} cell diffs (e.g. col={col!r} row={idx} A={va!r} B={vb!r})"
    return report


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python compare_xlsx.py A.xlsx B.xlsx", file=sys.stderr)
        sys.exit(2)
    a, b = Path(sys.argv[1]), Path(sys.argv[2])
    if not a.exists():
        print(f"A not found: {a}", file=sys.stderr)
        sys.exit(2)
    if not b.exists():
        print(f"B not found: {b}", file=sys.stderr)
        sys.exit(2)
    print(f"A: {a}")
    print(f"B: {b}")
    print()
    r = compare(a, b)
    bad = {k: v for k, v in r.items() if v != "OK"}
    for k, v in r.items():
        flag = "OK  " if v == "OK" else "DIFF"
        print(f"  [{flag}] {k}: {v}")
    print()
    print(f"Result: {len(r) - len(bad)} sheet(s) OK, {len(bad)} sheet(s) differ")
    sys.exit(0 if not bad else 1)
