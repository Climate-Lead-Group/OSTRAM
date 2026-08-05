"""Shared interconnector direction taxonomy and result aggregation."""

from __future__ import annotations

from typing import Iterable, Mapping

import pandas as pd

from ostram.pipeline.scenarios.rules.set_interconnector_direction import (
    TRN_PREFIX,
    VALID_DIRECTIONS,
    parse_tech_regions,
)


def interconnector_metadata(
    entries: Iterable[Mapping[str, object] | str],
) -> list[dict[str, object]]:
    """Validate manifest interconnectors using the shared TRN taxonomy."""

    normalized: list[dict[str, object]] = []
    seen: set[str] = set()
    for entry in entries:
        record = {"technology": entry} if isinstance(entry, str) else dict(entry)
        tech = str(record.get("technology", "")).strip()
        if not tech.startswith(TRN_PREFIX) or len(tech) != 13:
            raise ValueError(f"invalid interconnector technology: {tech!r}")
        if tech in seen:
            raise ValueError(f"duplicate interconnector metadata: {tech}")
        direction = str(record.get("direction", "bidirectional")).lower()
        if direction not in VALID_DIRECTIONS:
            raise ValueError(f"invalid direction {direction!r} for {tech}")
        source, destination = parse_tech_regions(tech)
        normalized.append(
            {
                **record,
                "technology": tech,
                "source_region": source,
                "destination_region": destination,
                "direction": direction,
            }
        )
        seen.add(tech)
    return normalized


def interconnector_series(
    frame: pd.DataFrame,
    technologies: Iterable[str],
) -> dict[str, dict[int, float]]:
    """Aggregate capacity/production for declared links in a compact fixture."""

    if "TECHNOLOGY" not in frame.columns or "YEAR" not in frame.columns:
        return {}
    selected = frame[frame["TECHNOLOGY"].astype(str).isin(set(technologies))]
    metrics = ("TotalCapacityAnnual", "ProductionByTechnologyAnnual")
    result: dict[str, dict[int, float]] = {}
    for metric in metrics:
        if metric not in selected.columns:
            continue
        values = pd.to_numeric(selected[metric], errors="coerce")
        grouped = values.groupby(pd.to_numeric(selected["YEAR"], errors="coerce")).sum()
        result[metric] = {
            int(year): float(value)
            for year, value in grouped.items()
            if pd.notna(year) and pd.notna(value)
        }
    return result
