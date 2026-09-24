#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
================================================================================
CAPITAL INVESTMENT ANNUALIZATION FOR OSTRAM
================================================================================
Author: Climate Lead Group, Andrey Salazar-Vargas

Purpose: Convert the lump-sum ``CapitalInvestment`` and
         ``CapitalInvestmentStorage`` results of the combined inputs/outputs
         CSV into Capital Recovery Factor (CRF) payment streams and append
         them to the same file as ``CapitalInvestmentAnnualized`` and
         ``CapitalInvestmentStorageAnnualized`` rows.

Method
------
Each investment made in year ``y`` for an asset with operational life ``n``
and discount rate ``r`` becomes ``n`` equal annual payments of
``investment * CRF(r, n)`` in years ``y .. y + n - 1`` (truncated at the
model horizon).  Payments of overlapping investments of the same asset are
accumulated.  The discounted sum of the payments equals the investment.

Both ``n`` and ``r`` are taken from the model the solver actually ran:
``OperationalLife`` / ``OperationalLifeStorage`` from the input parameters
present in the combined file, and ``DiscountRate`` / ``DiscountRateStorage``
from each scenario's compiled datafile (``<Scenario>_<Future>.txt``).  A
non-integer life (for example 27.5 years) is rounded half-up to a whole
number of payments so that the discounted payments still recover the
investment exactly.  No built-in numeric defaults are used: when a value
cannot be found and no explicit override was given, the run stops with an
error that names the missing asset, region or datafile.

Memory
------
The combined file is read in chunks and only the handful of columns needed
are kept, so the campaign-wide table is never loaded.  The existing rows are
streamed to a new file with two extra (empty) fields and the annualized rows
are appended at the end; the file is then replaced atomically.
================================================================================
"""

from __future__ import annotations

import argparse
import os
import re
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Sequence

import numpy as np
import pandas as pd


# ================================
# USER-CONFIGURABLE PARAMETERS
# ================================
INPUT_FILENAME = "OSTRAM_Combined_Inputs_Outputs.csv"  # default combined file for the CLI

CAPITAL_COLUMN = "CapitalInvestment"
NEW_COLUMN_NAME = "CapitalInvestmentAnnualized"

# Rows are read from the combined file in chunks of this many rows.
CHUNK_SIZE = 500_000


@dataclass(frozen=True)
class CapitalSeries:
    """One capital variable and the model parameters that annualize it."""

    capital_column: str
    life_column: str
    rate_parameter: str
    asset_column: str
    output_column: str

    @property
    def key_columns(self) -> list[str]:
        return ["Future", "Scenario", "REGION", self.asset_column]


CAPITAL_SERIES: tuple[CapitalSeries, ...] = (
    CapitalSeries(
        capital_column=CAPITAL_COLUMN,
        life_column="OperationalLife",
        rate_parameter="DiscountRate",
        asset_column="TECHNOLOGY",
        output_column=NEW_COLUMN_NAME,
    ),
    CapitalSeries(
        capital_column="CapitalInvestmentStorage",
        life_column="OperationalLifeStorage",
        rate_parameter="DiscountRateStorage",
        asset_column="STORAGE",
        output_column="CapitalInvestmentStorageAnnualized",
    ),
)


# ================================
# CAPITAL RECOVERY FACTOR
# ================================
def calculate_crf(discount_rate, lifetime):
    """
    Capital Recovery Factor: CRF = r (1 + r)^n / ((1 + r)^n - 1).

    Works element-wise on NumPy arrays.  A zero rate degenerates to 1 / n.
    """
    rate = np.asarray(discount_rate, dtype=float)
    life = np.asarray(lifetime, dtype=float)
    with np.errstate(divide="ignore", invalid="ignore"):
        growth = (1.0 + rate) ** life
        crf = np.where(rate == 0, 1.0 / life, rate * growth / (growth - 1.0))
    if crf.ndim == 0:
        return float(crf)
    return crf


# ================================
# DATAFILE PARAMETERS
# ================================
@dataclass(frozen=True)
class ParameterTable:
    """A GMPL ``param`` block: its default and its explicit rows."""

    default: float | None
    values: dict[tuple[str, ...], float]

    def lookup(self, key: tuple[str, ...]) -> float | None:
        return self.values.get(tuple(key), self.default)


_PARAM_HEADER = re.compile(
    r"^\s*param(?:\s+default\s+(?P<default>[-+0-9.eE]+))?\s*:\s*(?P<name>\w+)\s*:=\s*$"
)


def parse_datafile_parameter(datafile_path, parameter_name: str) -> ParameterTable:
    """Read one ``param [default X] : NAME :=`` block of a GMPL datafile."""
    default: float | None = None
    values: dict[tuple[str, ...], float] = {}
    inside = False
    with open(datafile_path, "r", encoding="utf-8", errors="replace") as handle:
        for line in handle:
            if not inside:
                match = _PARAM_HEADER.match(line)
                if match and match.group("name") == parameter_name:
                    inside = True
                    if match.group("default") is not None:
                        default = float(match.group("default"))
                continue
            stripped = line.strip()
            if stripped.startswith(";"):
                break
            if not stripped or stripped.startswith("#"):
                continue
            tokens = stripped.split()
            values[tuple(tokens[:-1])] = float(tokens[-1])
    return ParameterTable(default=default, values=values)


# ================================
# VECTORIZED ANNUALIZATION KERNEL
# ================================
def annualize_series(capex: pd.DataFrame, lifetimes: pd.DataFrame, rates: pd.DataFrame, horizon_end: int) -> pd.DataFrame:
    """
    Turn investments into accumulated CRF payment streams.

    Parameters
    ----------
    capex : DataFrame
        Key columns (ending with the asset column), ``YEAR`` and ``value``.
    lifetimes : DataFrame
        The same key columns plus ``life`` (years, may be non-integer).
    rates : DataFrame
        ``Future``, ``Scenario``, ``REGION`` and ``rate``.
    horizon_end : int
        Last model year; payments after it are dropped.

    Returns
    -------
    DataFrame
        Key columns, ``YEAR`` (int) and ``annualized`` (> 0), sorted by keys.
    """
    key_columns = [c for c in capex.columns if c not in ("YEAR", "value")]
    empty = pd.DataFrame(columns=key_columns + ["YEAR", "annualized"])
    capex = capex.dropna(subset=["value"])
    capex = capex[capex["value"] > 0]
    if capex.empty:
        return empty

    groups = capex[key_columns].drop_duplicates(ignore_index=True)
    groups = groups.merge(lifetimes[key_columns + ["life"]].drop_duplicates(key_columns), on=key_columns, how="left")
    missing_life = groups[groups["life"].isna()]
    if not missing_life.empty:
        raise ValueError(
            "Missing operational life for: "
            + ", ".join("/".join(map(str, row)) for row in missing_life[key_columns].itertuples(index=False))
        )
    rate_keys = ["Future", "Scenario", "REGION"]
    groups = groups.merge(rates[rate_keys + ["rate"]].drop_duplicates(rate_keys), on=rate_keys, how="left")
    missing_rate = groups[groups["rate"].isna()]
    if not missing_rate.empty:
        raise ValueError(
            "Missing discount rate for: "
            + ", ".join("/".join(map(str, row)) for row in missing_rate[rate_keys].drop_duplicates().itertuples(index=False))
        )

    life = np.floor(groups["life"].to_numpy(dtype=float) + 0.5).astype(int)
    if (life < 1).any():
        raise ValueError("Operational life must round to at least one year")
    crf = calculate_crf(groups["rate"].to_numpy(dtype=float), life)

    year_min = int(capex["YEAR"].min())
    horizon_end = int(horizon_end)
    if horizon_end < year_min:
        return empty
    years = np.arange(year_min, horizon_end + 1)
    n_groups, n_years = len(groups), len(years)

    # Investment matrix (groups x years).
    group_index = pd.MultiIndex.from_frame(groups[key_columns])
    row = group_index.get_indexer(pd.MultiIndex.from_frame(capex[key_columns]))
    col = capex["YEAR"].to_numpy(dtype=int) - year_min
    keep = col < n_years
    investment = np.zeros((n_groups, n_years))
    np.add.at(investment, (row[keep], col[keep]), capex["value"].to_numpy(dtype=float)[keep])

    # Payment in year t = CRF * sum of investments in (t - life, t].
    cumulative = np.cumsum(investment, axis=1)
    start = np.arange(n_years)[None, :] - life[:, None]
    earlier = np.take_along_axis(cumulative, np.clip(start, 0, n_years - 1), axis=1)
    window = cumulative - np.where(start >= 0, earlier, 0.0)
    annualized = window * crf[:, None]

    rows, cols = np.nonzero(annualized > 0)
    result = groups.loc[rows, key_columns].reset_index(drop=True)
    result["YEAR"] = years[cols]
    result["annualized"] = annualized[rows, cols]
    return result.sort_values(key_columns + ["YEAR"], kind="mergesort", ignore_index=True)


# ================================
# COMBINED FILE PROCESSING
# ================================
def _line_terminator(path: Path) -> str:
    with open(path, "rb") as handle:
        first = handle.readline()
    return "\r\n" if first.endswith(b"\r\n") else "\n"


def _datafile_path(executables_dir, future, scenario) -> Path:
    folder = f"{scenario}_{future}"
    return Path(executables_dir) / folder / f"{folder}.txt"


def _collect_rows(input_path: Path, header: list[str], series: Sequence[CapitalSeries]):
    """Stream the combined file keeping only capital/lifetime rows."""
    wanted = ["Future", "Scenario", "REGION", "YEAR"]
    wanted += [s.asset_column for s in series]
    wanted += [s.capital_column for s in series]
    wanted += [s.life_column for s in series if s.life_column in header]
    usecols = [c for c in header if c in wanted]
    value_columns = [s.capital_column for s in series] + [s.life_column for s in series if s.life_column in header]

    kept, rows_read, year_max = [], 0, None
    year_is_float = True
    for chunk in pd.read_csv(input_path, usecols=usecols, chunksize=CHUNK_SIZE, low_memory=False):
        rows_read += len(chunk)
        year_is_float = pd.api.types.is_float_dtype(chunk["YEAR"])
        chunk_max = chunk["YEAR"].max()
        if pd.notna(chunk_max):
            year_max = chunk_max if year_max is None else max(year_max, chunk_max)
        mask = chunk[value_columns].notna().any(axis=1)
        if mask.any():
            kept.append(chunk.loc[mask])
    rows = pd.concat(kept, ignore_index=True) if kept else pd.DataFrame(columns=usecols)
    return rows, rows_read, year_max, year_is_float


def _rates_for(capex: pd.DataFrame, item: CapitalSeries, executables_dir, discount_rate) -> pd.DataFrame:
    triples = capex[["Future", "Scenario", "REGION"]].drop_duplicates(ignore_index=True)
    if discount_rate is not None:
        triples["rate"] = float(discount_rate)
        return triples
    if executables_dir is None:
        raise FileNotFoundError(
            f"No executables folder given to read {item.rate_parameter}; pass discount_rate explicitly"
        )
    rates = []
    tables: dict[tuple, ParameterTable] = {}
    for future, scenario, region in triples.itertuples(index=False):
        key = (future, scenario)
        if key not in tables:
            datafile = _datafile_path(executables_dir, future, scenario)
            if not datafile.exists():
                raise FileNotFoundError(
                    f"Datafile {datafile} not found; cannot read {item.rate_parameter} "
                    f"for scenario {scenario} (pass discount_rate explicitly to override)"
                )
            tables[key] = parse_datafile_parameter(datafile, item.rate_parameter)
        rates.append(tables[key].lookup((str(region),)))
    triples["rate"] = rates
    return triples


def _lifetimes_for(capex: pd.DataFrame, rows: pd.DataFrame, item: CapitalSeries, executables_dir, asset_lifetime) -> pd.DataFrame:
    keys = item.key_columns
    groups = capex[keys].drop_duplicates(ignore_index=True)
    if asset_lifetime is not None:
        groups["life"] = float(asset_lifetime)
        return groups
    if item.life_column in rows.columns:
        explicit = rows.loc[rows[item.life_column].notna(), keys + [item.life_column]]
        explicit = explicit.rename(columns={item.life_column: "life"}).drop_duplicates(keys)
        groups = groups.merge(explicit, on=keys, how="left")
    else:
        groups["life"] = np.nan
    missing = groups["life"].isna()
    if missing.any() and executables_dir is not None:
        # Fall back to the datafile default the model itself solved with.
        defaults: dict[tuple, float | None] = {}
        for idx in groups.index[missing]:
            future, scenario = groups.at[idx, "Future"], groups.at[idx, "Scenario"]
            if (future, scenario) not in defaults:
                datafile = _datafile_path(executables_dir, future, scenario)
                defaults[(future, scenario)] = (
                    parse_datafile_parameter(datafile, item.life_column).default if datafile.exists() else None
                )
            groups.at[idx, "life"] = defaults[(future, scenario)]
    return groups


def _rewrite_with_appended(input_path: Path, added_columns: list[str], new_rows: pd.DataFrame, terminator: str) -> None:
    padding = ("," * len(added_columns)).encode("utf-8")
    term = terminator.encode("utf-8")
    tmp_path = input_path.with_name(input_path.name + ".tmp")
    with open(input_path, "rb") as source, open(tmp_path, "wb") as target:
        header = source.readline().rstrip(b"\r\n")
        target.write(header + ("," + ",".join(added_columns)).encode("utf-8") + term)
        for line in source:
            line = line.rstrip(b"\r\n")
            if line:
                target.write(line + padding + term)
        if not new_rows.empty:
            target.write(new_rows.to_csv(index=False, header=False, lineterminator=terminator).encode("utf-8"))
    os.replace(tmp_path, input_path)


def annualize_capital_investment(
    input_file_path=None,
    executables_dir=None,
    discount_rate=None,
    asset_lifetime=None,
    verbose=True,
):
    """
    Append annualized capital investment rows to a combined inputs/outputs CSV.

    Parameters
    ----------
    input_file_path : str or Path, optional
        Combined CSV to update in place. Defaults to ``INPUT_FILENAME``.
    executables_dir : str or Path, optional
        Folder with one ``<Scenario>_<Future>/<Scenario>_<Future>.txt`` datafile
        per scenario; used to read the discount rates (and the operational life
        default when an asset has no explicit value).
    discount_rate : float, optional
        Overrides the datafile discount rate for every series and region.
    asset_lifetime : float, optional
        Overrides the operational life for every asset.
    verbose : bool
        Print progress.

    Returns
    -------
    dict
        ``rows_read``, ``rows_appended``, ``columns_added``, ``horizon_end``.

    Raises
    ------
    FileNotFoundError
        Missing combined file or scenario datafile.
    ValueError
        File already annualized, or a lifetime/rate cannot be resolved.
    """
    input_path = Path(input_file_path or INPUT_FILENAME)
    if not input_path.exists():
        raise FileNotFoundError(f"Input file '{input_path.resolve()}' not found")

    def say(message):
        if verbose:
            print(message)

    say("=" * 60)
    say("CAPITAL INVESTMENT ANNUALIZATION")
    say("=" * 60)
    say(f"Start time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    say(f"Input file: {input_path}")

    header = list(pd.read_csv(input_path, nrows=0).columns)
    series = [s for s in CAPITAL_SERIES if s.capital_column in header]
    summary = {"rows_read": 0, "rows_appended": 0, "columns_added": [], "horizon_end": None}
    if not series:
        say("No capital investment columns found; nothing to annualize.")
        return summary
    already = [s.output_column for s in series if s.output_column in header]
    if already:
        raise ValueError(f"{input_path} already contains {', '.join(already)}; refusing to annualize twice")
    if "YEAR" not in header:
        raise ValueError("Column 'YEAR' not found in the CSV; it is required for the payment schedule")

    rows, rows_read, year_max, year_is_float = _collect_rows(input_path, header, series)
    summary["rows_read"] = rows_read
    if year_max is None or pd.isna(year_max):
        raise ValueError("No YEAR values found; cannot determine the model horizon")
    horizon_end = int(year_max)
    summary["horizon_end"] = horizon_end
    say(f"Rows read: {rows_read:,}; capital/lifetime rows kept: {len(rows):,}; horizon: {horizon_end}")

    added_columns = [s.output_column for s in series]
    output_columns = header + added_columns
    new_blocks = []
    for item in series:
        keys = item.key_columns
        capex = rows.loc[rows[item.capital_column].notna(), keys + ["YEAR", item.capital_column]]
        capex = capex.rename(columns={item.capital_column: "value"})
        if capex.empty:
            say(f"{item.capital_column}: no investments found")
            continue
        lifetimes = _lifetimes_for(capex, rows, item, executables_dir, asset_lifetime)
        rates = _rates_for(capex, item, executables_dir, discount_rate)
        result = annualize_series(capex, lifetimes, rates, horizon_end)
        say(
            f"{item.capital_column}: {len(capex):,} investments over {capex[keys].drop_duplicates().shape[0]:,} "
            f"assets -> {len(result):,} annualized rows in {item.output_column}"
        )
        block = result[keys].copy()
        block["YEAR"] = result["YEAR"].to_numpy(dtype=float if year_is_float else int)
        block[item.output_column] = result["annualized"].to_numpy()
        new_blocks.append(block)

    if new_blocks:
        new_rows = pd.concat(new_blocks, ignore_index=True, sort=False).reindex(columns=output_columns)
    else:
        new_rows = pd.DataFrame(columns=output_columns)
    _rewrite_with_appended(input_path, added_columns, new_rows, _line_terminator(input_path))

    summary["rows_appended"] = len(new_rows)
    summary["columns_added"] = added_columns
    say(f"Appended {len(new_rows):,} rows; added columns: {', '.join(added_columns)}")
    say(f"End time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    return summary


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Append annualized capital investment rows to a combined OSTRAM CSV.")
    parser.add_argument("--input", default=INPUT_FILENAME, help="combined inputs/outputs CSV (updated in place)")
    parser.add_argument("--executables", default=None, help="folder with <Scenario>_<Future>/<Scenario>_<Future>.txt datafiles")
    parser.add_argument("--discount-rate", type=float, default=None, help="override the model discount rate")
    parser.add_argument("--asset-lifetime", type=float, default=None, help="override the model operational life")
    parser.add_argument("--quiet", action="store_true")
    args = parser.parse_args(argv)
    try:
        annualize_capital_investment(
            input_file_path=args.input,
            executables_dir=args.executables,
            discount_rate=args.discount_rate,
            asset_lifetime=args.asset_lifetime,
            verbose=not args.quiet,
        )
    except (FileNotFoundError, ValueError) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
