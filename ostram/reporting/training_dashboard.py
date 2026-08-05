"""Generate a compact, self-contained dashboard from synthetic or real results."""

from __future__ import annotations

from datetime import datetime, timezone
import html
import json
from pathlib import Path
from typing import Iterable, Mapping

import pandas as pd

from .interconnector_direction import interconnector_metadata, interconnector_series


GENERATION_PREFIXES = {
    "PWRSPV": "Solar PV",
    "PWRWON": "Onshore wind",
    "PWRWOF": "Offshore wind",
    "PWRHYD": "Hydro",
    "PWRCOA": "Coal",
    "PWRNGS": "Gas",
    "PWRNUC": "Nuclear",
    "PWRBIO": "Biomass",
}
STORAGE_PREFIXES = ("PWRSDS", "PWRLDS")
VRE_PREFIXES = ("PWRSPV", "PWRWON", "PWRWOF")


def tech_family(technology: object) -> str | None:
    value = str(technology)
    for prefix, family in GENERATION_PREFIXES.items():
        if value.startswith(prefix):
            return family
    return None


def _series(frame: pd.DataFrame, column: str) -> dict[int, float]:
    if column not in frame.columns or "YEAR" not in frame.columns:
        return {}
    years = pd.to_numeric(frame["YEAR"], errors="coerce")
    values = pd.to_numeric(frame[column], errors="coerce")
    grouped = values.groupby(years).sum(min_count=1)
    return {
        int(year): float(value)
        for year, value in grouped.items()
        if pd.notna(year) and pd.notna(value)
    }


def _region_filter(frame: pd.DataFrame, region: str | None) -> pd.DataFrame:
    if region is None or region == "System":
        return frame
    masks = []
    if "REGION" in frame.columns:
        masks.append(frame["REGION"].astype(str) == region)
    if "TECHNOLOGY" in frame.columns:
        masks.append(frame["TECHNOLOGY"].astype(str).str.endswith(region))
    if not masks:
        return frame.iloc[0:0]
    mask = masks[0]
    for extra in masks[1:]:
        mask = mask | extra
    return frame[mask]


def aggregate_metrics(
    frame: pd.DataFrame,
    *,
    scenario: str,
    region: str | None,
    interconnectors: Iterable[Mapping[str, object]] = (),
) -> dict[str, object]:
    """Aggregate only declared columns; missing metrics remain empty."""

    selected = frame
    if "Scenario" in selected.columns:
        selected = selected[selected["Scenario"].astype(str) == scenario]
    selected = _region_filter(selected, region)
    tech = selected.get("TECHNOLOGY", pd.Series("", index=selected.index)).astype(str)
    generation: dict[str, dict[int, float]] = {}
    capacity: dict[str, dict[int, float]] = {}
    for family in GENERATION_PREFIXES.values():
        family_mask = tech.map(tech_family) == family
        family_frame = selected[family_mask]
        generation[family] = _series(family_frame, "ProductionByTechnologyAnnual")
        capacity[family] = _series(family_frame, "TotalCapacityAnnual")
    storage = selected[tech.str.startswith(STORAGE_PREFIXES)]
    generation_rows = selected[tech.map(tech_family).notna()]
    vre_rows = selected[tech.str.startswith(VRE_PREFIXES)]
    total_generation = _series(generation_rows, "ProductionByTechnologyAnnual")
    vre_generation = _series(vre_rows, "ProductionByTechnologyAnnual")
    vre_share = {
        year: (vre_generation.get(year, 0.0) / value if value else 0.0)
        for year, value in total_generation.items()
    }
    costs = {
        name: float(pd.to_numeric(selected[name], errors="coerce").sum())
        for name in ("TotalDiscountedCost", "CapitalInvestment")
        if name in selected.columns
    }
    declared = interconnector_metadata(interconnectors)
    return {
        "available": not selected.empty,
        "generation": generation,
        "capacity": capacity,
        "emissions": _series(selected, "AnnualEmissions"),
        "storage": _series(storage, "TotalCapacityAnnual"),
        "vre_share": vre_share,
        "cost": costs,
        "interconnectors": interconnector_series(
            selected, [item["technology"] for item in declared]
        ),
    }


def build_dashboard_data(
    snapshots: Iterable[tuple[str, Path]],
    *,
    profile_id: str,
    manifest: Path,
    workspace: Path,
    metadata: Mapping[str, object],
) -> dict[str, object]:
    country_regions_raw = metadata.get("country_regions", [])
    if isinstance(country_regions_raw, Mapping):
        country_regions = [
            {"region": str(region), "label": str(label)}
            for region, label in country_regions_raw.items()
        ]
    elif isinstance(country_regions_raw, list):
        country_regions = country_regions_raw
    else:
        raise ValueError("profile metadata country_regions must be a list or mapping")
    regions: list[str] = []
    for item in country_regions:
        region = item.get("region") if isinstance(item, Mapping) else item
        if not isinstance(region, str) or not region.strip():
            raise ValueError(f"invalid country region metadata: {item!r}")
        regions.append(region.strip())
    interconnectors = metadata.get("interconnectors", [])
    if not isinstance(interconnectors, list):
        raise ValueError("profile metadata interconnectors must be a list")
    scenario_hint = metadata.get("scenarios", [])
    data: dict[str, object] = {}
    for label, path in snapshots:
        frame = pd.read_csv(path, low_memory=False)
        scenarios = (
            sorted(frame["Scenario"].dropna().astype(str).unique())
            if "Scenario" in frame.columns
            else [str(item) for item in scenario_hint]
        )
        data[label] = {
            scenario: {
                region: aggregate_metrics(
                    frame,
                    scenario=scenario,
                    region=region,
                    interconnectors=interconnectors,
                )
                for region in ["System", *regions]
            }
            for scenario in scenarios
        }
    return {
        "schema": "ostram-profile-report-v1",
        "generated_at_utc": datetime.now(timezone.utc).isoformat(),
        "profile_id": profile_id,
        "manifest": str(Path(manifest).resolve()),
        "workspace": str(Path(workspace).resolve()),
        "country_regions": country_regions,
        "interconnectors": interconnector_metadata(interconnectors),
        "effective_values": dict(metadata.get("effective_values", {})),
        "snapshots": data,
    }


def render_html(data: Mapping[str, object]) -> str:
    payload = json.dumps(data, ensure_ascii=False, sort_keys=True).replace("</", "<\\/")
    title = html.escape(f"OSTRAM {data['profile_id']} profile report")
    return f"""<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width">
<title>{title}</title><style>
body{{font:14px system-ui,sans-serif;max-width:1000px;margin:2rem auto;padding:0 1rem;color:#173c3d}}
pre{{white-space:pre-wrap;background:#f3f7f5;padding:1rem;border-radius:8px}} .meta{{color:#526b6b}}
</style></head><body><h1>{title}</h1>
<p class="meta">Manifest: {html.escape(str(data['manifest']))}<br>Workspace: {html.escape(str(data['workspace']))}</p>
<div id="report"></div><script id="ostram-profile-data" type="application/json">{payload}</script>
<script>const d=JSON.parse(document.getElementById('ostram-profile-data').textContent);
document.getElementById('report').innerHTML='<h2>Snapshots</h2><pre>'+JSON.stringify(d.snapshots,null,2)+'</pre>';</script>
</body></html>"""


def generate_report(
    snapshots: Iterable[tuple[str, Path]],
    output: Path,
    **context,
) -> Path:
    data = build_dashboard_data(snapshots, **context)
    output = Path(output).resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(render_html(data), encoding="utf-8", newline="\n")
    return output
