"""
Parse DAYPART_DEF out of both scripts and verify they are in sync.

Used as a gate before every generation run: the generator does NOT error on a
fabric mismatch against the Ninja config, it silently remaps CFs
(INPUT_CONTRACT.md section 5), so the only defence is checking the source of
truth in both files.

Exit 0 = in sync and structurally valid; exit 1 = not.
"""
import ast
import os
import sys

MODULE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
GEN = os.path.join(MODULE, 'scripts', 'build_ostram_timeslices.py')
REB = os.path.join(MODULE, 'scripts', 'rebuild_reninja_timeslices_latest.py')


def parse_daypart_def(path):
    """Pull the module-level DAYPART_DEF assignment out via AST, no import."""
    with open(path, encoding='utf-8') as f:
        tree = ast.parse(f.read())
    for node in tree.body:
        if isinstance(node, ast.Assign):
            for tgt in node.targets:
                if isinstance(tgt, ast.Name) and tgt.id == 'DAYPART_DEF':
                    return [tuple(x) for x in ast.literal_eval(node.value)], node.lineno
    raise SystemExit(f"DAYPART_DEF not found in {path}")


def structural_check(dp, label):
    """Same invariants both scripts enforce at import time."""
    errs = []
    if dp[0][2] != 0:
        errs.append(f"first daypart starts at {dp[0][2]}, not 0")
    if dp[-1][3] != 24:
        errs.append(f"last daypart ends at {dp[-1][3]}, not 24")
    for i in range(1, len(dp)):
        if dp[i][2] != dp[i - 1][3]:
            errs.append(f"gap/overlap: {dp[i-1][0]} ends {dp[i-1][3]}, "
                        f"{dp[i][0]} starts {dp[i][2]}")
    hours = sum(d[3] - d[2] for d in dp)
    if hours != 24:
        errs.append(f"daypart hours sum to {hours}, not 24")
    codes = [d[0] for d in dp]
    if len(set(codes)) != len(codes):
        errs.append(f"duplicate daypart codes: {codes}")
    return errs, hours


def main():
    gen_dp, gen_line = parse_daypart_def(GEN)
    reb_dp, reb_line = parse_daypart_def(REB)

    print("=" * 72)
    print("DAYPART_DEF SYNC CHECK")
    print("=" * 72)
    print(f"\ngenerator  {os.path.basename(GEN)}  (line {gen_line})")
    for d in gen_dp:
        print(f"    {d[0]}  {d[1]:<14s} {d[2]:02d}-{d[3]:02d}  ({d[3]-d[2]} h)")
    print(f"\nrebuilder  {os.path.basename(REB)}  (line {reb_line})")
    for d in reb_dp:
        print(f"    {d[0]}  {d[1]:<14s} {d[2]:02d}-{d[3]:02d}  ({d[3]-d[2]} h)")

    # Boundaries are what the generator's runtime match test compares
    # (code, start, end) — labels are display-only and excluded.
    gen_bounds = [(d[0], d[2], d[3]) for d in gen_dp]
    reb_bounds = [(d[0], d[2], d[3]) for d in reb_dp]
    bounds_match = gen_bounds == reb_bounds
    labels_match = gen_dp == reb_dp

    errs, hours = structural_check(gen_dp, 'generator')
    n_ts = 4 * len(gen_dp)

    print(f"\n{'-'*72}")
    print(f"  boundaries in sync : {bounds_match}")
    print(f"  labels also match  : {labels_match}")
    print(f"  hours covered      : {hours} (contiguous 0->24: {not errs})")
    print(f"  dayparts           : {len(gen_dp)}")
    print(f"  timeslices         : 4 seasons x {len(gen_dp)} dp = {n_ts}")
    if errs:
        print("  structural errors  :")
        for e in errs:
            print(f"      - {e}")
    print("-" * 72)

    ok = bounds_match and not errs
    print(f"\n{'IN SYNC: True' if ok else 'IN SYNC: False'}")
    return 0 if ok else 1


if __name__ == '__main__':
    sys.exit(main())
