"""Base-year build lock and its two pre-solve checks.

In the lock years (2023 to 2026 in the tracked configuration), every technology
of the GENERATION family except the backstop ``PWRBCK*``, of the two STORAGE
families and of the INTERCONNECTORS family may add only the capacity its plant
list gives: ``TotalAnnualMaxCapacityInvestment`` equals the
``TotalAnnualMinCapacityInvestment`` the model receives for the same cell, and
0 where none is given. Where the scenario already holds a lower ceiling, the
lower value stays. The internal grid, fuel supply and the backstop keep their
own rules.

B2 applies the lock to the normalised CSVs of one scenario, before the otoole
conversion, so the solver datafile and every later adjustment start from it.
The careful reserve-margin repair skips the locked cells.

Two pre-solve checks run with the lock:

(a) For each locked technology with an annual activity lower limit in a lock
    year, the most the locked fleet can generate is its residual capacity plus
    its cumulative locked builds, times ``CapacityToActivityUnit``, times the
    sum over time slices of ``CapacityFactor`` x ``YearSplit``, times the
    ``AvailabilityFactor`` where it is below 1. A lower limit above that value
    cannot be met; it is lowered to that value and reported as a plant-list gap.
(b) For each historical trade lower bound (``D4_I1_TRADE_*_Lower`` in the model
    file) in a lock year of a scenario where the trade bounds are active, the
    links it names must be able to carry it with their residual capacity plus
    their locked builds. Where they cannot, those links are left out of the
    lock, so their investment cells in the lock years stay as they are.
"""
from __future__ import annotations

import csv
import re
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd

from ostram.paths import resolve_paths


LOCKED_FAMILIES = ("GENERATION", "STORAGE_LONG", "STORAGE_SHORT", "INTERCONNECTORS")
UNLOCKED_PREFIXES = ("PWRBCK",)

MAX_INV = "TotalAnnualMaxCapacityInvestment"
MIN_INV = "TotalAnnualMinCapacityInvestment"
LOWER = "TotalTechnologyAnnualActivityLowerLimit"

# otoole defaults (ostram/resources/compilation/conversion_format.yaml).
DEFAULTS = {
    MAX_INV: -1.0,
    MIN_INV: 0.0,
    "ResidualCapacity": 0.0,
    "CapacityFactor": 1.0,
    "AvailabilityFactor": 1.0,
    "CapacityToActivityUnit": 1.0,
    "OperationalLife": 1.0,
    LOWER: 0.0,
    "InputActivityRatio": 0.0,
    "OutputActivityRatio": 0.0,
}

_TRADE_LOWER = re.compile(r"^s\.t\. (D4_I1_TRADE_\w+_Lower)\{[^}]*\}: (.*) >= ([-+0-9.eE]+);\s*$")
_TRADE_TERM = re.compile(
    r'RateOfActivity\[r,l,"(\w+)",(\d+),(\d{4})\]\*YearSplit\[l,\d{4}\]'
    r'(?:\*(Output|Input)ActivityRatio\[r,"\w+","(\w+)",\d+,\d{4}\])?'
)


def lock_years(params: dict) -> range | None:
    """Return the lock years of the B2 configuration, or None when the lock is off."""
    if not params.get("base_year_lock_active", False):
        return None
    first = int(params["base_year_lock_first_year"])
    last = int(params["base_year_lock_last_year"])
    if first > last:
        raise ValueError(f"base_year_lock_first_year {first} is after base_year_lock_last_year {last}")
    return range(first, last + 1)


def locked_technologies(taxonomy: Path | None = None) -> dict[str, str]:
    """Return {technology: family} for the locked families of the taxonomy."""
    path = Path(taxonomy) if taxonomy is not None else resolve_paths().interconnector_taxonomy
    out = {}
    with path.open(encoding="utf-8-sig", newline="") as handle:
        for row in csv.reader(handle):
            if len(row) < 2 or not row[0].strip() or not row[1].strip():
                continue
            family, tech = row[0].strip(), row[1].strip()
            if family in LOCKED_FAMILIES and not tech.startswith(UNLOCKED_PREFIXES):
                out[tech] = family
    return out


@dataclass(frozen=True)
class TradeBound:
    """One historical trade lower bound of the model file."""

    name: str
    year: int
    terms: tuple  # (technology, mode, ratio kind or "", fuel or "")
    rhs: float


def trade_lower_bounds(model_text: str) -> list[TradeBound]:
    """Return the ``D4_I1_TRADE_*_Lower`` constraints of the model file text."""
    bounds = []
    for line in model_text.splitlines():
        match = _TRADE_LOWER.match(line)
        if not match:
            continue
        terms = _TRADE_TERM.findall(match.group(2))
        if not terms:
            continue
        years = {int(term[2]) for term in terms}
        if len(years) != 1:
            raise ValueError(f"{match.group(1)} spans more than one year: {sorted(years)}")
        bounds.append(TradeBound(
            name=match.group(1),
            year=years.pop(),
            terms=tuple((t, int(m), kind, fuel) for t, m, _, kind, fuel in terms),
            rhs=float(match.group(3)),
        ))
    return bounds


@dataclass
class LockReport:
    """What the lock and its checks did for one scenario."""

    scenario: str
    cells: int = 0
    changed: list = field(default_factory=list)       # (tech, year, before, after)
    tighter_kept: list = field(default_factory=list)  # (tech, year, ceiling, plant list)
    gaps: list = field(default_factory=list)          # (tech, year, observed, fleet)
    exempt: list = field(default_factory=list)        # (constraint, year, links, bound, capacity)

    def lines(self) -> list[str]:
        out = [f"[base-year-lock] {self.scenario}: {self.cells} locked cells, "
               f"{len(self.changed)} changed, {len(self.tighter_kept)} lower ceilings kept, "
               f"{len(self.gaps)} plant-list gaps lowered, {len(self.exempt)} trade bounds exempted"]
        out += [f"[base-year-lock] lower ceiling kept: {t} {y} ceiling {c!r} plant list {p!r}"
                for t, y, c, p in self.tighter_kept]
        out += [f"[base-year-lock] plant-list gap: {t} {y} observed lower limit {o!r} PJ, "
                f"locked fleet {f!r} PJ; lower limit lowered to the fleet value"
                for t, y, o, f in self.gaps]
        out += [f"[base-year-lock] trade bound {n} ({y}) needs {b!r} PJ, links {'+'.join(links)} "
                f"carry {c!r} PJ with residual and locked builds; their lock-year investment cells "
                f"are left as they are" for n, y, links, b, c in self.exempt]
        return out


def _read(folder: Path, name: str) -> pd.DataFrame | None:
    path = folder / f"{name}.csv"
    if not path.is_file():
        return None
    return pd.read_csv(path)


def _series(df: pd.DataFrame | None, keys: list[str]) -> dict:
    # A row with an empty value is absent: otoole writes no value for it, so the
    # model receives the parameter's default.
    if df is None or df.empty:
        return {}
    return {tuple(str(v) for v in row[:-1]): float(row[-1])
            for row in df[keys + ["VALUE"]].itertuples(index=False, name=None) if not pd.isna(row[-1])}


def _write_sorted(folder: Path, name: str, df: pd.DataFrame) -> None:
    # Same order as the normalising step (all columns).
    df.sort_values(by=list(df.columns)).to_csv(folder / f"{name}.csv", index=False)


def apply_base_year_lock(
    folder: str | Path,
    scenario_name: str,
    years: range,
    *,
    trade_active: bool,
    model_text: str,
    taxonomy: Path | None = None,
) -> LockReport:
    """Apply the lock and its two checks to the normalised CSVs in ``folder``.

    Rewrites ``TotalAnnualMaxCapacityInvestment.csv`` and, for plant-list gaps,
    ``TotalTechnologyAnnualActivityLowerLimit.csv`` in place.
    """
    folder = Path(folder)
    report = LockReport(scenario_name)
    techs_df = _read(folder, "TECHNOLOGY")
    present = set(techs_df["VALUE"].astype(str)) if techs_df is not None else set()
    locked = {t: f for t, f in locked_technologies(taxonomy).items() if t in present}
    ys = [str(y) for y in years]

    max_df = _read(folder, MAX_INV)
    min_inv = _series(_read(folder, MIN_INV), ["TECHNOLOGY", "YEAR"])
    residual = _series(_read(folder, "ResidualCapacity"), ["TECHNOLOGY", "YEAR"])
    life = _series(_read(folder, "OperationalLife"), ["TECHNOLOGY"])
    cta = _series(_read(folder, "CapacityToActivityUnit"), ["TECHNOLOGY"])
    avail = _series(_read(folder, "AvailabilityFactor"), ["TECHNOLOGY", "YEAR"])
    ysplit = _series(_read(folder, "YearSplit"), ["TIMESLICE", "YEAR"])
    cf = _series(_read(folder, "CapacityFactor"), ["TECHNOLOGY", "TIMESLICE", "YEAR"])
    slices = defaultdict(list)
    for (ts, y), v in ysplit.items():
        slices[y].append((ts, v))

    def g(series, name, *key):
        return series.get(tuple(key), DEFAULTS[name])

    def activity_per_gw(t, y):
        total = sum(cf.get((t, ts, y), DEFAULTS["CapacityFactor"]) * v for ts, v in slices[y])
        return g(cta, "CapacityToActivityUnit", t) * total * min(g(avail, "AvailabilityFactor", t, y), 1.0)

    # The locked value of each cell: the plant list, or a lower ceiling already there.
    current = _series(max_df, ["TECHNOLOGY", "YEAR"])
    lock = {}
    for t in sorted(locked):
        for y in ys:
            plant = g(min_inv, MIN_INV, t, y)
            ceiling = current.get((t, y), DEFAULTS[MAX_INV])
            lock[(t, y)] = ceiling if ceiling != -1 and ceiling < plant else plant

    def fleet(t, y):
        span = g(life, "OperationalLife", t)
        cum = sum(lock.get((t, w), g(min_inv, MIN_INV, t, w)) for w in ys
                  if int(w) <= int(y) and int(y) - int(w) < span)
        return g(residual, "ResidualCapacity", t, y) + cum

    # Check (b): trade lower bounds the locked links cannot carry. Runs first,
    # because a link it leaves out of the lock keeps its cells.
    exempt = set()
    if trade_active:
        ratios = {}
        for name in ("InputActivityRatio", "OutputActivityRatio"):
            ratios[name.replace("ActivityRatio", "")] = _series(
                _read(folder, name), ["TECHNOLOGY", "FUEL", "MODE_OF_OPERATION", "YEAR"])
        for bound in trade_lower_bounds(model_text):
            y = str(bound.year)
            if y not in ys or bound.rhs <= 0:
                continue
            by_tech = defaultdict(float)
            for t, m, kind, fuel in bound.terms:
                ratio = ratios[kind].get((t, fuel, str(m), y), 0.0) if kind else 1.0
                by_tech[t] = max(by_tech[t], ratio)
            capacity = sum(fleet(t, y) * activity_per_gw(t, y) * r for t, r in by_tech.items())
            if bound.rhs > capacity:
                links = tuple(sorted(by_tech))
                exempt |= set(links)
                report.exempt.append((bound.name, bound.year, links, bound.rhs, capacity))

    # The lock.
    written = {k: v for k, v in lock.items() if k[0] not in exempt}
    for (t, y), value in sorted(written.items()):
        report.cells += 1
        if value < g(min_inv, MIN_INV, t, y):
            report.tighter_kept.append((t, int(y), value, g(min_inv, MIN_INV, t, y)))
        if current.get((t, y)) != value:
            report.changed.append((t, int(y), current.get((t, y)), value))
    if written:
        region = str(max_df["REGION"].iloc[0]) if max_df is not None and not max_df.empty else "GLOBAL"
        rows = pd.DataFrame([(region, t, int(y), v) for (t, y), v in sorted(written.items())],
                            columns=["REGION", "TECHNOLOGY", "YEAR", "VALUE"])
        if max_df is not None and not max_df.empty:
            others = [(str(t), str(y)) not in written for t, y in zip(max_df["TECHNOLOGY"], max_df["YEAR"])]
            rows = pd.concat([max_df[others], rows], ignore_index=True)
        _write_sorted(folder, MAX_INV, rows)

    # Check (a): activity lower limits the locked fleet cannot reach.
    lower_df = _read(folder, LOWER)
    if lower_df is not None and not lower_df.empty:
        lowered = False
        values = lower_df["VALUE"].astype(float).tolist()
        for i, (t, y) in enumerate(zip(lower_df["TECHNOLOGY"].astype(str), lower_df["YEAR"].astype(str))):
            if t not in locked or t in exempt or y not in ys:
                continue
            most = fleet(t, y) * activity_per_gw(t, y)
            if values[i] > most:
                report.gaps.append((t, int(y), values[i], most))
                values[i] = most
                lowered = True
        if lowered:
            lower_df = lower_df.copy()
            lower_df["VALUE"] = values
            _write_sorted(folder, LOWER, lower_df)
    return report


def locked_flow_keys(technologies, years: range | None, taxonomy: Path | None = None) -> set:
    """Return the (technology, year) cells the lock holds, for the given technologies.

    The reserve-margin repair uses it to skip locked cells. Interconnectors left
    out by check (b) are not repair targets, so the taxonomy alone decides.
    """
    if years is None:
        return set()
    locked = locked_technologies(taxonomy)
    return {(t, y) for t in technologies if t in locked for y in years}
