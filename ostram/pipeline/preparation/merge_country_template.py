"""Merge a reviewed country template into the active preparation workspace."""

from __future__ import annotations

import argparse
from datetime import datetime, timezone
import os
from pathlib import Path
import shutil
from typing import Any, Sequence

import pandas as pd

from ostram.paths import resolve_paths
from ostram.profiles import (
    DEFAULT_PROFILE,
    active_profile_id,
)


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


def _default_input_dir(paths) -> Path:
    """Keep full's legacy staging target; profiles merge their mutable authority."""

    if active_profile_id() == DEFAULT_PROFILE:
        return paths.preparation_workspace / "og_csvs_inputs"
    return paths.osemosys_inputs


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
        required = ["region", "lat", "long"]
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


SNAPSHOT_PREFIX = "_post_a2_snapshot_"


def _invalidate_stale_snapshots(a1_outputs: Path) -> list[str]:
    """Rename existing post-A2 snapshots so A1/A2 are forced to regenerate.

    After merging a new country, the old snapshots don't include the new
    country's technologies and must be rebuilt.  Snapshots are renamed
    (not deleted) so data is never lost.
    """
    if not a1_outputs.is_dir():
        return []
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    renamed: list[str] = []
    for entry in sorted(a1_outputs.iterdir()):
        if entry.is_dir() and entry.name.startswith(SNAPSHOT_PREFIX):
            new_name = entry.with_name(f"{entry.name}_stale_{stamp}")
            entry.rename(new_name)
            renamed.append(entry.name)
    return renamed


def _find_a1_outputs(paths) -> Path:
    """Resolve the A1_Outputs directory robustly across profile contexts.

    ``paths.a1_outputs`` applies ``.resolve()`` which can behave differently
    on Windows depending on the profile environment.  This helper tries
    multiple strategies to locate the directory that actually contains the
    post-A2 snapshots.
    """
    # Strategy 1: the canonical property
    candidate = paths.a1_outputs
    if candidate.is_dir() and any(
        p.is_dir() and p.name.startswith(SNAPSHOT_PREFIX)
        for p in candidate.iterdir()
    ):
        return candidate
    # Strategy 2: preparation_workspace / A1_Outputs (no .resolve())
    candidate = paths.preparation_workspace / "A1_Outputs"
    if candidate.is_dir() and any(
        p.is_dir() and p.name.startswith(SNAPSHOT_PREFIX)
        for p in candidate.iterdir()
    ):
        return candidate
    # Strategy 3: walk up from workspace looking for preparation/A1_Outputs
    for parent in [paths.workspace, paths.workspace.parent]:
        for prep in parent.rglob("preparation"):
            candidate = prep / "A1_Outputs"
            if candidate.is_dir() and any(
                p.is_dir() and p.name.startswith(SNAPSHOT_PREFIX)
                for p in candidate.iterdir()
            ):
                return candidate
    # Fallback: return paths.a1_outputs even if we couldn't verify it
    return paths.a1_outputs


def main(argv: Sequence[str] | None = None) -> int:
    paths = resolve_paths()
    parser = argparse.ArgumentParser(prog="python -m ostram country merge")
    parser.add_argument(
        "country",
        nargs="?",
        help="ISO-3 country whose generated workspace template should be merged",
    )
    parser.add_argument("--template", type=Path)
    parser.add_argument("--input-dir", type=Path, default=_default_input_dir(paths))
    parser.add_argument("--centerpoints", type=Path, default=paths.preparation_inputs / "centerpoints.csv")
    args = parser.parse_args(argv)
    if args.country and args.template:
        parser.error("country and --template are mutually exclusive")
    if args.country:
        country = args.country.strip().upper()
        if len(country) != 3 or not country.isalpha():
            parser.error(f"country must be an ISO-3 code: {args.country!r}")
        template = paths.preparation_workspace / "country_templates" / country
    elif args.template:
        template = args.template
    else:
        parser.error("provide COUNTRY or --template PATH")
    counts = merge_country_template(template, args.input_dir, args.centerpoints)
    print(f"Merged country template: {counts}")

    # For non-default profiles the primary merge target is the authority
    # directory, but A1 reads from og_csvs_inputs.  Mirror the merge there
    # so the next A1 run picks up the new country data.
    og_csvs = paths.preparation_workspace / "og_csvs_inputs"
    if og_csvs.is_dir() and og_csvs.resolve() != args.input_dir.resolve():
        og_counts = merge_country_template(template, og_csvs, args.centerpoints)
        print(f"Also merged into og_csvs_inputs: {og_counts}")

    # Invalidate stale post-A2 snapshots so the next run regenerates A1/A2
    # with the newly merged country data.
    a1_dir = _find_a1_outputs(paths)
    print(f"Checking for stale snapshots in: {a1_dir}")
    stale = _invalidate_stale_snapshots(a1_dir)
    if stale:
        print(
            f"Invalidated {len(stale)} stale post-A2 snapshot(s): "
            f"{', '.join(stale)}"
        )
        print(
            "The next 'ostram run' will regenerate A1 and A2 to include "
            "the merged country."
        )
    else:
        print("No stale post-A2 snapshots found to invalidate.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
