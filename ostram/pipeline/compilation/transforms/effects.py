"""Injectable workbook, CSV, configuration, and pickle effects for B1."""

from __future__ import annotations

from collections.abc import Callable, Mapping
from typing import Any
import pickle

import pandas as pd
import yaml


def read_config(
    path: Any,
    *,
    opener: Callable[..., Any] | None = None,
    loader: Callable[[Any], Any] | None = None,
) -> Any:
    """Read a YAML config with the predecessor's text-mode/default-encoding call."""
    if opener is None:
        opener = open
    if loader is None:
        loader = yaml.safe_load
    with opener(path, "r") as stream:
        return loader(stream)


def open_workbook(
    path: Any, *, factory: Callable[[Any], Any] | None = None
) -> Any:
    if factory is None:
        factory = pd.ExcelFile
    return factory(path)


def read_csv(path: Any, *, reader: Callable[[Any], Any] | None = None) -> Any:
    if reader is None:
        reader = pd.read_csv
    return reader(path)


def load_pickle(
    path: Any,
    *,
    opener: Callable[..., Any] | None = None,
    loader: Callable[[Any], Any] | None = None,
) -> Any:
    """Load a pickle while preserving the predecessor's non-context-managed open."""
    if opener is None:
        opener = open
    if loader is None:
        loader = pickle.load
    return loader(opener(path, "rb"))


def _write_frame_to_excel(
    frame: pd.DataFrame,
    writer: Any,
    sheet_name: Any,
    write_frame: Callable[..., Any] | None,
) -> None:
    if write_frame is None:
        frame.to_excel(writer, sheet_name=sheet_name, index=False)
    else:
        write_frame(frame, writer, sheet_name=sheet_name, index=False)


def write_completed_demand_workbook(
    path: Any,
    frame: pd.DataFrame,
    initial_year: Any,
    sheet_name: Any,
    *,
    writer_factory: Callable[..., Any] | None = None,
    write_frame: Callable[..., Any] | None = None,
) -> pd.DataFrame:
    """Write the completed demand workbook and return its rounded frame.

    There is deliberately no context manager or failure cleanup: a conversion or
    write failure propagates before ``close`` exactly as it did in the predecessor.
    """
    if writer_factory is None:
        writer_factory = pd.ExcelWriter
    writer = writer_factory(path, engine="xlsxwriter")
    frame[initial_year] = frame[initial_year].astype(float)
    rounded_frame = frame.round(4)
    _write_frame_to_excel(rounded_frame, writer, sheet_name, write_frame)
    writer.close()
    return rounded_frame


def write_sheet_mapping_workbook(
    path: Any,
    sheets: Mapping[Any, pd.DataFrame],
    *,
    writer_factory: Callable[..., Any] | None = None,
    write_frame: Callable[..., Any] | None = None,
) -> None:
    """Write mapping values in insertion order, rounding each temporary frame."""
    if writer_factory is None:
        writer_factory = pd.ExcelWriter
    writer = writer_factory(path, engine="xlsxwriter")
    sheet_names = list(sheets.keys())
    for index in range(len(sheet_names)):
        sheet_name = sheet_names[index]
        frame = sheets[sheet_name].round(4)
        _write_frame_to_excel(frame, writer, sheet_name, write_frame)
    writer.close()


def write_structure_workbook(
    path: Any,
    frame: pd.DataFrame,
    sheet_name: Any,
    *,
    writer_factory: Callable[..., Any] | None = None,
    write_frame: Callable[..., Any] | None = None,
) -> None:
    if writer_factory is None:
        writer_factory = pd.ExcelWriter
    writer = writer_factory(path, engine="xlsxwriter")
    _write_frame_to_excel(frame, writer, sheet_name, write_frame)
    writer.close()
