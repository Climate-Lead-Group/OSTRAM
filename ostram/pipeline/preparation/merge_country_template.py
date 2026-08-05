"""Merge a reviewed country template into the active preparation workspace."""

from __future__ import annotations

import argparse
from datetime import datetime, timezone
from pathlib import Path
import shutil
from typing import Sequence

import pandas as pd

from ostram.paths import resolve_paths


DATASETS = (
    "TECHNOLOGY", "FUEL", "EMISSION", "STORAGE",
    "CapitalCost", "FixedCost", "VariableCost", "ResidualCapacity",
    "CapacityFactor", "AvailabilityFactor", "InputActivityRatio",
    "OutputActivityRatio", "EmissionActivityRatio", "SpecifiedAnnualDemand",
    "SpecifiedDemandProfile", "OperationalLife", "CapacityToActivityUnit",
    "TotalAnnualMaxCapacity", "TotalAnnualMaxCapacityInvestment",
    "TotalTechnologyAnnualActivityUpperLimit", "ReserveMarginTagTechnology",
    "ReserveMarginTagFuel", "CapitalCostStorage", "OperationalLifeStorage",
    "StorageLevelStart", "ResidualStorageCapacity", "TechnologyToStorage",
    "TechnologyFromStorage",
)


def merge_country_template(
    template_dir: Path,
    input_dir: Path,
    centerpoints_path: Path,
) -> dict[str, int]:
    """Merge exact-schema CSVs, backing up each existing destination once."""

    template_dir = Path(template_dir).resolve()
    input_dir = Path(input_dir).resolve()
    centerpoints_path = Path(centerpoints_path).resolve()
    if not template_dir.is_dir():
        raise FileNotFoundError(f"country template directory not found: {template_dir}")
    input_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    counts: dict[str, int] = {}
    for name in DATASETS:
        source = template_dir / f"{name}.csv"
        if not source.is_file():
            continue
        incoming = pd.read_csv(source)
        if incoming.empty:
            continue
        destination = input_dir / source.name
        if destination.is_file():
            existing = pd.read_csv(destination)
            if list(existing.columns) != list(incoming.columns):
                raise ValueError(
                    f"schema mismatch for {name}: {list(existing.columns)} != "
                    f"{list(incoming.columns)}"
                )
            shutil.copy2(destination, destination.with_suffix(f".{stamp}.bak.csv"))
            merged = pd.concat([existing, incoming], ignore_index=True)
            merged = merged.drop_duplicates().reset_index(drop=True)
        else:
            merged = incoming.drop_duplicates().reset_index(drop=True)
        merged.to_csv(destination, index=False)
        counts[name] = len(merged)

    source_centerpoint = template_dir / "centerpoint.csv"
    if source_centerpoint.is_file():
        incoming = pd.read_csv(source_centerpoint)
        required = ["region", "latitude", "longitude"]
        if list(incoming.columns) != required:
            raise ValueError(
                f"centerpoint schema must be {required}, got {list(incoming.columns)}"
            )
        centerpoints_path.parent.mkdir(parents=True, exist_ok=True)
        if centerpoints_path.is_file():
            existing = pd.read_csv(centerpoints_path)
            if list(existing.columns) != required:
                raise ValueError("existing centerpoints schema differs from template")
            shutil.copy2(
                centerpoints_path,
                centerpoints_path.with_suffix(f".{stamp}.bak.csv"),
            )
            regions = set(incoming["region"].astype(str))
            existing = existing[~existing["region"].astype(str).isin(regions)]
            incoming = pd.concat([existing, incoming], ignore_index=True)
        incoming.sort_values("region").to_csv(centerpoints_path, index=False)
        counts["centerpoint"] = len(incoming)
    return counts


def main(argv: Sequence[str] | None = None) -> int:
    paths = resolve_paths()
    parser = argparse.ArgumentParser(prog="python -m ostram country merge")
    parser.add_argument("--template", type=Path, required=True)
    parser.add_argument("--input-dir", type=Path, default=paths.preparation_workspace / "og_csvs_inputs")
    parser.add_argument("--centerpoints", type=Path, default=paths.preparation_inputs / "centerpoints.csv")
    args = parser.parse_args(argv)
    counts = merge_country_template(args.template, args.input_dir, args.centerpoints)
    print(f"Merged country template: {counts}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
