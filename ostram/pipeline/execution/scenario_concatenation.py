"""Streaming cross-scenario concatenation for the B2 final post-processing.

The legacy implementation loaded every scenario's inputs and outputs into
memory, concatenated them into three campaign-wide DataFrames and sorted
each one.  With 16 scenarios that meant three copies of a 7.6 million row
table and a ``MemoryError`` on a 32 GB machine.

This module produces the same three files (inputs, outputs, combined) by
processing one scenario at a time and appending its sorted block to the
CSV files.  Because the global sort order starts with ``Future`` and
``Scenario``, and the accumulated minimum-investment column is grouped by
scenario, writing scenario blocks in sorted order is equivalent to sorting
the whole campaign at once.  Peak memory is therefore bounded by the
largest single scenario instead of by the campaign.
"""

from __future__ import annotations

import os
import shutil
from dataclasses import dataclass
from datetime import date
from typing import Iterable

import numpy as np
import pandas as pd


# Set columns moved to the front of every file, in this fixed order.
KEY_SET_COLUMNS = [
    "REGION", "YEAR", "TECHNOLOGY", "FUEL", "EMISSION", "MODE_OF_OPERATION",
    "TIMESLICE", "STORAGE", "SEASON", "DAYTYPE", "DAILYTIMEBRACKET",
]
METADATA_COLUMNS = ["Future", "Scenario"]
SORT_COLUMNS = ["Future", "Scenario", "REGION", "TECHNOLOGY", "YEAR"]

MIN_INVESTMENT_COLUMN = "TotalAnnualMinCapacityInvestment"
ACCUMULATED_MIN_INVESTMENT_COLUMN = "AccumulatedTotalAnnualMinCapacityInvestment"
ACCUMULATION_GROUP = ["Future", "Scenario", "TECHNOLOGY"]

SKIPPED_ENTRIES = {"default", "__pycache__", "local_dataset_creator_0.py"}


def active_output_csv_candidates(params, scenario_future_name):
    """
    Return output CSV names in the same suffix order used by main_executer.

    The solver/otoole path can become, for example:
      Pre_processed_BAU_0_NoStorage_OpenBCK_RMCarefulXLSX_output.csv

    The final scenario concatenator used to look only for:
      Pre_processed_BAU_0_Output.csv

    Keep the active chained name first, with legacy fallbacks after it.
    """
    base = f"{params['preprocess_data_name']}{scenario_future_name}"
    chain_parts = []

    if params.get("storage_delay_active", False):
        chain_parts.append(params.get("storage_delay_suffix", "StorageDelayN5"))
    if params.get("strip_storage_active", False):
        chain_parts.append(params.get("strip_storage_suffix", "NoStorage"))
    if params.get("open_pwrbck_active", False):
        chain_parts.append(params.get("open_pwrbck_suffix", "OpenBCK"))
    if params.get("reserve_margin_repair_active", False):
        chain_parts.append(params.get("reserve_margin_repair_suffix", "RMRepair"))
    if params.get("reserve_margin_xlsx_active", False):
        chain_parts.append(params.get("reserve_margin_xlsx_suffix", "RMCarefulXLSX"))

    candidates = []
    if chain_parts:
        candidates.append(f"{base}_{'_'.join(chain_parts)}{params['output_files']}.csv")

    candidates.extend([
        f"{base}{params['output_files']}.csv",
        f"{base}_Output.csv",
    ])

    return candidates


@dataclass(frozen=True)
class ScenarioFiles:
    """Per-scenario CSV locations discovered under the executables folder."""

    name: str
    scenario: str
    future: str
    input_path: str | None
    output_path: str | None


def discover_scenario_files(here, params) -> list[ScenarioFiles]:
    """List scenario folders in sorted order with their input/output CSVs."""
    base_input_path = os.path.join(here, params["executables"])
    if not os.path.isdir(base_input_path):
        return []

    found: list[ScenarioFiles] = []
    for scenario_future_name in sorted(os.listdir(base_input_path)):
        if scenario_future_name.lower() in SKIPPED_ENTRIES:
            continue
        scenario_path = os.path.join(base_input_path, scenario_future_name)
        if not os.path.isdir(scenario_path):
            continue

        scenario, _, future = scenario_future_name.rpartition("_")
        if not scenario:
            scenario, future = scenario_future_name, ""

        input_path = os.path.join(scenario_path, f"{scenario_future_name}_Input.csv")
        if not os.path.exists(input_path):
            input_path = None

        output_path = None
        for output_name in active_output_csv_candidates(params, scenario_future_name):
            candidate = os.path.join(scenario_path, output_name)
            if os.path.exists(candidate):
                output_path = candidate
                break

        if input_path is None and output_path is None:
            continue
        found.append(ScenarioFiles(scenario_future_name, scenario, future, input_path, output_path))
    return found


def ordered_columns(columns: Iterable[str]) -> list[str]:
    """Metadata first, then the known set columns, then the rest alphabetically."""
    present = set(columns)
    front = METADATA_COLUMNS + [c for c in KEY_SET_COLUMNS if c in present]
    rest = sorted(c for c in present if c not in front)
    return front + rest


def _header(path: str) -> list[str]:
    return list(pd.read_csv(path, nrows=0).columns)


def _read_block(path: str, scenario: str, future: str) -> pd.DataFrame:
    frame = pd.read_csv(path, low_memory=False)
    frame.insert(0, "Future", future)
    frame.insert(1, "Scenario", scenario)
    return frame


def _sorted(frame: pd.DataFrame, by: list[str]) -> pd.DataFrame:
    keys = [c for c in by if c in frame.columns]
    if not keys:
        return frame
    return frame.sort_values(by=keys, kind="mergesort", ignore_index=True)


def _accumulate_min_investment(frame: pd.DataFrame) -> pd.DataFrame:
    """Add the cumulative minimum-investment column, grouped within scenario."""
    frame[ACCUMULATED_MIN_INVESTMENT_COLUMN] = np.nan
    group_cols = [c for c in ACCUMULATION_GROUP if c in frame.columns]
    frame = _sorted(frame, group_cols + ["YEAR"])
    mask = frame[MIN_INVESTMENT_COLUMN].notna()
    if group_cols:
        accumulated = (
            frame.loc[mask]
            .groupby(group_cols, sort=False)[MIN_INVESTMENT_COLUMN]
            .cumsum()
        )
    else:
        accumulated = frame.loc[mask, MIN_INVESTMENT_COLUMN].cumsum()
    frame.loc[mask, ACCUMULATED_MIN_INVESTMENT_COLUMN] = accumulated
    return frame


class _CsvAppender:
    """Append DataFrame blocks with a fixed column layout to one CSV file."""

    def __init__(self, path: str, columns: list[str]):
        self.path = path
        self.columns = columns
        self._started = False

    def append(self, block: pd.DataFrame) -> None:
        block = block.reindex(columns=self.columns)
        block.to_csv(
            self.path,
            index=False,
            mode="a" if self._started else "w",
            header=not self._started,
        )
        self._started = True


def _dated_copy(path: str, today: str) -> None:
    shutil.copyfile(path, path.replace(".csv", f"_{today}.csv"))


def concatenate_all_scenarios(HERE, params):
    """
    Iterate over all scenario folders in ``params['executables']`` (excluding
    'Default'), read *_Input.csv and *_output.csv files, add scenario metadata
    columns and concatenate them into single CSV files for inputs, outputs and
    combined, one scenario block at a time.

    Args:
        HERE (str): Execution workspace that contains the executables folder.
        params (dict):
          - executables (str): Path to the base directory containing the scenario folders.
          - prefix_final_files (str): Prefix for the final file names.
          - inputs_file (str): Base name for the inputs CSV.
          - outputs_file (str): Base name for the outputs CSV.
          - combined_file (str, optional): Base name for the combined inputs+outputs CSV.
    Returns:
        tuple: (input_csv_path, output_csv_path, combined_csv_path); an entry is
        ``None`` when no scenario contributed rows to that file.
    """
    scenarios = discover_scenario_files(HERE, params)
    input_files = [s for s in scenarios if s.input_path]
    output_files = [s for s in scenarios if s.output_path]

    # Pass 1: headers only, to fix the column layout before writing anything.
    input_columns: set[str] = set(METADATA_COLUMNS)
    output_columns: set[str] = set(METADATA_COLUMNS)
    for item in input_files:
        input_columns.update(_header(item.input_path))
    for item in output_files:
        output_columns.update(_header(item.output_path))

    cols_in = ordered_columns(input_columns) if input_files else []
    cols_out = ordered_columns(output_columns) if output_files else []
    cols_comb = ordered_columns(input_columns | output_columns) if (input_files and output_files) else []
    accumulate = bool(cols_comb) and MIN_INVESTMENT_COLUMN in cols_comb
    if accumulate:
        cols_comb = cols_comb + [ACCUMULATED_MIN_INVESTMENT_COLUMN]

    prefix = os.path.join(HERE, params["prefix_final_files"])
    combined_name = params.get("combined_file", "Combined_Inputs_Outputs.csv")
    path_in = prefix + params["inputs_file"] if cols_in else None
    path_out = prefix + params["outputs_file"] if cols_out else None
    path_comb = prefix + combined_name if cols_comb else None

    writer_in = _CsvAppender(path_in, cols_in) if path_in else None
    writer_out = _CsvAppender(path_out, cols_out) if path_out else None
    writer_comb = _CsvAppender(path_comb, cols_comb) if path_comb else None

    # Pass 2: one scenario at a time, in sorted order.
    for item in scenarios:
        df_in = _read_block(item.input_path, item.scenario, item.future) if item.input_path else None
        df_out = _read_block(item.output_path, item.scenario, item.future) if item.output_path else None

        if writer_in is not None and df_in is not None:
            writer_in.append(_sorted(df_in, SORT_COLUMNS))
        if writer_out is not None and df_out is not None:
            writer_out.append(_sorted(df_out, SORT_COLUMNS))
        if writer_comb is not None:
            parts = [f for f in (df_in, df_out) if f is not None]
            block = pd.concat(parts, ignore_index=True, sort=True)
            del parts
            block = _sorted(block, SORT_COLUMNS)
            if accumulate:
                block = _accumulate_min_investment(block)
            writer_comb.append(block)
            del block
        del df_in, df_out

    today = date.today().isoformat()
    for path in (path_in, path_out):
        if path:
            _dated_copy(path, today)
    # The combined dated copy is created by the orchestrator, after the
    # optional annualization step.

    return path_in, path_out, path_comb
