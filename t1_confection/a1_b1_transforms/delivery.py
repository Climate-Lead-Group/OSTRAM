"""Output delivery helpers preserving the legacy B1 write sequence."""

from __future__ import annotations

from collections.abc import Callable, Iterable, Mapping
from typing import Any
import os

import pandas as pd


def clean_parameter_tables(
    main_tables: Mapping[Any, pd.DataFrame],
) -> tuple[dict[Any, pd.DataFrame], dict[Any, pd.DataFrame]]:
    """Remove missing keys and recreate both delivery mappings from main tables.

    The second mapping intentionally does *not* use the predecessor's accumulated
    NDP mapping.  That historical quirk is part of accepted output behavior and is
    made explicit here rather than corrected during the structural refactor.
    """
    cleaned_main = {
        key: value for key, value in main_tables.items() if not pd.isna(key)
    }
    cleaned_additional = {
        key: value for key, value in cleaned_main.items() if not pd.isna(key)
    }
    return cleaned_main, cleaned_additional


def _write_csv(
    frame: pd.DataFrame,
    path: Any,
    writer: Callable[..., Any] | None,
) -> None:
    if writer is None:
        frame.to_csv(path, index=False, header=True)
    else:
        writer(frame, path, index=False, header=True)


def deliver_main_csvs(
    output_root: Any,
    main_scenario: Any,
    parameter_tables: Mapping[Any, pd.DataFrame],
    structure_values: Mapping[str, list[Any]],
    *,
    makedirs: Callable[..., Any] | None = None,
    dataframe_factory: Callable[..., pd.DataFrame] | None = None,
    csv_writer: Callable[..., Any] | None = None,
) -> str:
    """Write main parameters in insertion order, then sorted structure files."""
    if makedirs is None:
        makedirs = os.makedirs
    if dataframe_factory is None:
        dataframe_factory = pd.DataFrame

    output_path = os.path.join(output_root, main_scenario)
    makedirs(output_path, exist_ok=True)

    table_names = list(parameter_tables.keys())
    for index in range(len(table_names)):
        name = table_names[index]
        frame = parameter_tables[name]
        _write_csv(
            frame,
            os.path.join(output_path, name + ".csv"),
            csv_writer,
        )

    for column_name, data_list in sorted(structure_values.items()):
        frame = dataframe_factory({"VALUE": data_list})
        file_path = os.path.join(output_path, f"{column_name}.csv")
        _write_csv(frame, file_path, csv_writer)
    return output_path


def deliver_additional_csvs(
    output_root: Any,
    scenarios: Iterable[Any],
    main_scenario: Any,
    parameter_tables: Mapping[Any, pd.DataFrame],
    structure_values: Mapping[str, list[Any]],
    *,
    makedirs: Callable[..., Any] | None = None,
    dataframe_factory: Callable[..., pd.DataFrame] | None = None,
    csv_writer: Callable[..., Any] | None = None,
) -> None:
    """Write additional scenarios in configured order, including duplicates."""
    if makedirs is None:
        makedirs = os.makedirs
    if dataframe_factory is None:
        dataframe_factory = pd.DataFrame

    for scenario in scenarios:
        output_path = os.path.join(output_root, scenario)
        makedirs(output_path, exist_ok=True)

        for name, frame in sorted(parameter_tables.items()):
            scenario_frame = frame.replace(
                {"Scenario": {main_scenario: scenario}}
            )
            _write_csv(
                scenario_frame,
                os.path.join(output_path, f"{name}.csv"),
                csv_writer,
            )

        for column_name, data_list in sorted(structure_values.items()):
            structure_frame = dataframe_factory({"VALUE": data_list})
            _write_csv(
                structure_frame,
                os.path.join(output_path, f"{column_name}.csv"),
                csv_writer,
            )
