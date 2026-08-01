"""Combine otoole result CSV files into one scenario output table."""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Sequence

import pandas as pd


SET_COLUMNS = (
    "YEAR",
    "TECHNOLOGY",
    "TIMESLICE",
    "FUEL",
    "EMISSION",
    "MODE_OF_OPERATION",
    "REGION",
    "SEASON",
    "DAYTYPE",
    "DAILYTIMEBRACKET",
    "STORAGE",
)


def concatenate_outputs(outputs_folder: Path, output_file: Path) -> Path | None:
    """Write ``output_file.csv`` from the non-empty CSVs in ``outputs_folder``."""

    if not outputs_folder.is_dir():
        return None

    frames: list[pd.DataFrame] = []
    parameters: list[str] = []
    allowed = {"Parameter", "VALUE", *SET_COLUMNS}
    for path in sorted(outputs_folder.iterdir(), key=lambda item: item.name):
        if not path.is_file():
            continue
        frame = pd.read_csv(path)
        frame = frame[[column for column in frame.columns if column in allowed]]
        frame["Parameter"] = path.stem
        if not frame.empty:
            frames.append(frame.dropna(axis=1, how="all"))
            parameters.append(path.stem)

    if not frames:
        return None

    combined = pd.concat(frames, ignore_index=True, sort=True)
    columns = sorted(set(combined.columns) & allowed)
    combined = combined[columns]
    first_parameter = parameters[0]
    merged = combined[combined["Parameter"] == first_parameter]
    merged = merged.rename(columns={"VALUE": first_parameter}).drop("Parameter", axis=1)
    merged = merged.assign(
        **{column: "nan" for column in SET_COLUMNS if column not in merged.columns}
    )

    for parameter in parameters[1:]:
        frame = combined[combined["Parameter"] == parameter]
        if frame.empty:
            continue
        frame = frame.rename(columns={"VALUE": parameter}).drop("Parameter", axis=1)
        frame = frame.assign(
            **{column: "nan" for column in SET_COLUMNS if column not in frame.columns}
        )
        merged = pd.merge(merged, frame, on=list(SET_COLUMNS), how="outer")

    destination = output_file.with_suffix(".csv")
    merged.to_csv(destination)
    return destination


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("outputs_folder", type=Path)
    parser.add_argument("output_file", type=Path)
    args = parser.parse_args(argv)
    concatenate_outputs(args.outputs_folder.resolve(), args.output_file.resolve())
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
