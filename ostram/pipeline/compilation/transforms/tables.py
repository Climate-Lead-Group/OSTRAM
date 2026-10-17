"""Pure tabular transformations used by the B1 compiler."""

from __future__ import annotations

from typing import Any, Mapping, NamedTuple, Sequence

import numpy as np
import pandas as pd


def normalize_year_like_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize Excel year headers to plain strings such as ``"2023"``.

    This intentionally retains the predecessor's identity behavior: if no header
    qualifies, the input object itself is returned.  Collisions are not rejected,
    so duplicate normalized column labels remain possible.
    """
    rename_map = {}
    for col in df.columns:
        if isinstance(col, (int, np.integer)):
            rename_map[col] = str(col)
        elif isinstance(col, float) and col.is_integer():
            rename_map[col] = str(int(col))
        elif isinstance(col, str):
            stripped = col.strip()
            if stripped.isdigit():
                rename_map[col] = str(int(stripped))
    if rename_map:
        df = df.rename(columns=rename_map)
    return df


def build_system_parameter_rows(
    system_parameters: pd.DataFrame,
    years: Sequence[int],
    run_settings: Mapping[str, Any],
) -> pd.DataFrame:
    """Expand non-missing system-parameter cells into long-form rows.

    Rows, duplicate keys, input index order, Python ``float`` conversion, and
    four-decimal rounding are intentionally left exactly as in the predecessor.
    Missing columns and invalid values therefore continue to raise their native
    ``KeyError`` and ``ValueError`` before a caller updates output dictionaries.
    """
    system_parameters = normalize_year_like_columns(system_parameters)
    accumulated_rows = []
    for row_index in system_parameters.index:
        parameter = system_parameters.loc[row_index, "Parameter"]
        for year_index in range(len(years)):
            year_key = str(years[year_index])
            value = system_parameters.loc[row_index, year_key]
            if pd.notna(value):
                accumulated_rows.append(
                    {
                        "PARAMETER": parameter,
                        "Scenario": run_settings["Main_Scenario"],
                        "REGION": run_settings["Region"],
                        "YEAR": years[year_index],
                        "Value": round(float(value), 4),
                    }
                )
    return pd.DataFrame(accumulated_rows)


class StructureTables(NamedTuple):
    """The wide structure workbook table and per-set CSV value lists."""

    table: pd.DataFrame
    values: dict[str, list[Any]]


def build_structure_tables(
    time_range_vector: list[int],
    all_tech_list: list[Any],
    all_fuel_list: list[Any],
    emissions_list: list[Any],
    run_settings: Mapping[str, Any],
    params: Mapping[str, Any],
) -> StructureTables:
    """Build the predecessor's padded structure table without sorting or deduping."""
    lengths = [
        len(time_range_vector),
        len(all_tech_list),
        len(run_settings["Timeslices"]),
        len(all_fuel_list),
        len(emissions_list),
        len(run_settings["Mode_of_Operation"]),
        len([run_settings["Region"]]),
        len(run_settings["Season"]),
        len(run_settings["DayType"]),
        len(run_settings["DailyTimeBracket"]),
        len(run_settings["Storage"]),
    ]
    maximum_length = max(lengths)

    structure_year = time_range_vector + [
        "" for _ in range(maximum_length - lengths[0])
    ]
    structure_technology = all_tech_list + [
        "" for _ in range(maximum_length - lengths[1])
    ]
    structure_timeslice = run_settings["Timeslices"] + [
        "" for _ in range(maximum_length - lengths[2])
    ]
    structure_fuel = all_fuel_list + [
        "" for _ in range(maximum_length - lengths[3])
    ]
    structure_emission = emissions_list + [
        "" for _ in range(maximum_length - lengths[4])
    ]
    structure_mode = run_settings["Mode_of_Operation"] + [
        "" for _ in range(maximum_length - lengths[5])
    ]
    structure_region = [run_settings["Region"]] + [
        "" for _ in range(maximum_length - lengths[6])
    ]
    structure_season = run_settings["Season"] + [
        "" for _ in range(maximum_length - lengths[7])
    ]
    structure_daytype = run_settings["DayType"] + [
        "" for _ in range(maximum_length - lengths[8])
    ]
    structure_daily_time_bracket = run_settings["DailyTimeBracket"] + [
        "" for _ in range(maximum_length - lengths[9])
    ]
    structure_storage = run_settings["Storage"] + [
        "" for _ in range(maximum_length - lengths[10])
    ]

    values = {
        "YEAR": structure_year,
        "TECHNOLOGY": structure_technology,
        "TIMESLICE": structure_timeslice,
        "FUEL": structure_fuel,
        "EMISSION": structure_emission,
        "MODE_OF_OPERATION": structure_mode,
        "REGION": structure_region,
        "DAYTYPE": structure_daytype,
        "DAILYTIMEBRACKET": structure_daily_time_bracket,
        "SEASON": structure_season,
        "STORAGE": structure_storage,
    }
    # Pandas 3 may infer its nullable StringDtype for string-only columns.
    # The compiler's supported contract is the predecessor's object dtype,
    # which also keeps mixed integer/blank padding stable across versions.
    table = pd.DataFrame(
        {
            name: pd.Series(values[name], dtype=object)
            for name in params["columns4"]
        },
        columns=params["columns4"],
    )
    return StructureTables(table=table, values=values)
