"""Profile-aware OSTRAM training dashboard.

Aggregates combined input/output result snapshots into one self-contained,
offline, interactive HTML report. This module ports the interactive renderer
and the proven metric semantics of the original UNESCAP training dashboard
into the profile/capture architecture:

* scenarios come from the snapshot data (or profile metadata hints);
* country regions, labels, storage families, interconnectors and effective
  values come from profile metadata — nothing model-specific is hardcoded;
* snapshots are the explicit ``(label, path)`` pairs handed in by the caller —
  the module never discovers inputs from the working directory.

The generated HTML embeds the ``ostram-profile-report-v1`` payload in a stable
``ostram-profile-data`` element and draws every chart with inline SVG and
vanilla JavaScript: no CDN, external script, font or network request.
"""

from __future__ import annotations

from datetime import datetime, timezone
import html as html_text
import json
from pathlib import Path
from typing import Iterable, Mapping

import pandas as pd

from ostram.pipeline.preparation.scenario_country_sync import technology_regions

from .interconnector_direction import interconnector_metadata


# Columns the aggregation understands. Snapshots are loaded through a header
# probe + usecols so absent optional columns simply yield empty metrics.
NEEDED_COLUMNS = (
    "Scenario", "REGION", "YEAR", "TECHNOLOGY", "EMISSION",
    "ProductionByTechnologyAnnual", "TotalCapacityAnnual",
    "NewCapacity", "TotalAnnualMaxCapacityInvestment",
    "AnnualEmissions", "TotalDiscountedCost", "CapitalInvestment",
)

# A scenario counts as solved when any of these solver outputs is present.
# Combined snapshots accumulate input rows for every scenario, so presence of
# rows alone does not mean the scenario has results.
OUTPUT_COLUMNS = (
    "ProductionByTechnologyAnnual", "TotalCapacityAnnual", "NewCapacity",
    "AnnualEmissions", "TotalDiscountedCost", "CapitalInvestment",
)

# Technology family taxonomy (labels + colors) shared with the historical
# scenario-analysis tooling for visual consistency. This is the OSTRAM-wide
# technology naming taxonomy, not a profile-specific routing table.
TECH_FAMILIES = {
    "PWRCOA": ("Coal",          "#4a4a4a"),
    "PWRNGS": ("Gas",           "#d4a017"),
    "PWROIL": ("Oil",           "#8b4513"),
    "PWRPET": ("Petroleum",     "#a0522d"),
    "PWRURN": ("Nuclear",       "#e74c3c"),
    "PWRHYD": ("Hydro",         "#2980b9"),
    "PWRSHP": ("Small Hydro",   "#5dade2"),
    "PWRSPV": ("Solar PV",      "#f39c12"),
    "PWRWON": ("Wind Onshore",  "#27ae60"),
    "PWRWOF": ("Wind Offshore", "#1abc9c"),
    "PWRCSP": ("CSP",           "#e67e22"),
    "PWRBIO": ("Biomass",       "#7d3c98"),
    "PWRWAS": ("Waste",         "#95a5a6"),
    "PWRGEO": ("Geothermal",    "#c0392b"),
    "PWROTH": ("Other",         "#bdc3c7"),
    "PWRCCS": ("CCS",           "#566573"),
}
ORDERED_FAMILIES = [
    "Coal", "Petroleum", "Oil", "Gas", "Nuclear", "Other", "CCS", "Waste",
    "Biomass", "Geothermal", "Hydro", "Small Hydro", "CSP", "Solar PV",
    "Wind Onshore", "Wind Offshore",
]

VRE_PREFIXES = ("PWRSPV", "PWRWON", "PWRWOF")
INTERNAL_TRANSMISSION_PREFIX = "PWRTRN"
BACKSTOP_PREFIX = "PWRBCK"
DEFAULT_STORAGE_PREFIXES = ("PWRSDS", "PWRLDS")
DEFAULT_YEAR_RANGE = {"start": 2023, "end": 2050}


def tech_family(technology: object) -> str | None:
    """Map a technology code to its family label, or None for non-generation."""
    value = str(technology)
    for prefix, (label, _color) in TECH_FAMILIES.items():
        if value.startswith(prefix):
            return label
    return None


def tech_color(family: str) -> str:
    """Hex color for a family label; grey fallback."""
    for _prefix, (label, color) in TECH_FAMILIES.items():
        if label == family:
            return color
    return "#cccccc"


def scenario_label(scenario: str) -> str:
    """Derive a short display label from a scenario id (no hardcoded names)."""
    parts = str(scenario).split("_")
    if len(parts) > 1 and 1 <= len(parts[0]) <= 2:
        return parts[0] + " · " + " ".join(parts[1:])
    return str(scenario)


def storage_prefixes(metadata: Mapping[str, object] | None) -> tuple[str, ...]:
    """Derive PWR storage-tech prefixes from profile storage metadata."""
    declared = (metadata or {}).get("storage", [])
    prefixes: list[str] = []
    if isinstance(declared, (list, tuple)):
        for entry in declared:
            code = str(entry).strip().upper()[:3]
            if len(code) == 3 and code.isalpha():
                prefix = "PWR" + code
                if prefix not in prefixes:
                    prefixes.append(prefix)
    return tuple(prefixes) if prefixes else DEFAULT_STORAGE_PREFIXES


def is_storage_tech(technology: object, storage: tuple[str, ...]) -> bool:
    return str(technology).startswith(storage)


def is_generation_tech(technology: object, storage: tuple[str, ...]) -> bool:
    """PWR* generation, excluding internal transmission, backstop and storage."""
    value = str(technology)
    if not value.startswith("PWR"):
        return False
    if value.startswith((INTERNAL_TRANSMISSION_PREFIX, BACKSTOP_PREFIX)):
        return False
    return not value.startswith(storage)


def is_vre_tech(technology: object) -> bool:
    return str(technology).startswith(VRE_PREFIXES)


def emission_serves_region(emission: object, region: str) -> bool:
    """Structural emission-to-region attribution (code suffix vs region code)."""
    code = str(emission).strip().upper()
    target = str(region).strip().upper()
    if not code or not target:
        return False
    return code.endswith(target) or code.endswith(target[:3])


def _load_frame(path: Path) -> pd.DataFrame:
    """Load only the columns the dashboard needs; coerce YEAR to Int64."""
    header = pd.read_csv(path, nrows=0)
    usable = [column for column in NEEDED_COLUMNS if column in header.columns]
    frame = pd.read_csv(path, usecols=usable or None, low_memory=False)
    if "YEAR" in frame.columns:
        frame["YEAR"] = pd.to_numeric(frame["YEAR"], errors="coerce").astype("Int64")
    return frame


def _dedup(frame: pd.DataFrame, columns: Iterable[str]) -> pd.DataFrame:
    """Select the identity+value columns and drop duplicated combined rows."""
    subset = [column for column in columns if column in frame.columns]
    return frame[subset].drop_duplicates()


def _scenario_frame(frame: pd.DataFrame, scenario: str) -> pd.DataFrame:
    if "Scenario" in frame.columns:
        return frame[frame["Scenario"].astype(str) == str(scenario)]
    return frame


def scenario_has_results(scenario_frame: pd.DataFrame) -> bool:
    """True when any solver output value exists for the scenario rows."""
    for column in OUTPUT_COLUMNS:
        if column in scenario_frame.columns:
            values = pd.to_numeric(scenario_frame[column], errors="coerce")
            if values.notna().any():
                return True
    return False


def _region_rows(frame: pd.DataFrame, region: str | None) -> pd.DataFrame:
    """Keep rows structurally attributable to the region (System keeps all)."""
    if region is None or region == "System":
        return frame
    masks = []
    if "REGION" in frame.columns:
        masks.append(frame["REGION"].astype(str) == region)
    if "TECHNOLOGY" in frame.columns:
        masks.append(
            frame["TECHNOLOGY"]
            .map(lambda tech: region in technology_regions(tech))
            .astype(bool)
        )
    if not masks:
        return frame.iloc[0:0]
    mask = masks[0]
    for extra in masks[1:]:
        mask = mask | extra
    return frame[mask]


def _stacked_by_family(
    scenario_frame: pd.DataFrame,
    value_column: str,
    region: str | None,
    storage: tuple[str, ...],
) -> dict[int, dict[str, float]]:
    """{year: {family: value}} over generation technologies for one region."""
    required = {value_column, "YEAR", "TECHNOLOGY"}
    if not required.issubset(scenario_frame.columns):
        return {}
    sub = scenario_frame[
        scenario_frame[value_column].notna()
        & scenario_frame["YEAR"].notna()
        & scenario_frame["TECHNOLOGY"].notna()
    ]
    sub = sub[sub["TECHNOLOGY"].map(lambda tech: is_generation_tech(tech, storage)).astype(bool)]
    sub = _region_rows(sub, region)
    sub = _dedup(sub, ("YEAR", "TECHNOLOGY", value_column)).copy()
    if sub.empty:
        return {}
    sub["family"] = sub["TECHNOLOGY"].map(tech_family)
    sub = sub[sub["family"].notna()]
    out: dict[int, dict[str, float]] = {}
    grouped = sub.groupby(["YEAR", "family"])[value_column].sum()
    for (year, family), value in grouped.items():
        out.setdefault(int(year), {})[family] = round(float(value), 4)
    return out


def emissions_series(
    scenario_frame: pd.DataFrame, region: str | None
) -> dict[int, float]:
    """Annual emissions summed by year; region attribution by EMISSION code.

    Values are intentionally not rounded so System stays the exact sum of the
    per-region series.
    """
    required = {"AnnualEmissions", "EMISSION", "YEAR"}
    if not required.issubset(scenario_frame.columns):
        return {}
    sub = scenario_frame[
        scenario_frame["AnnualEmissions"].notna() & scenario_frame["YEAR"].notna()
    ]
    if region is not None and region != "System":
        sub = sub[
            sub["EMISSION"].map(lambda code: emission_serves_region(code, region)).astype(bool)
        ]
    sub = _dedup(sub, ("YEAR", "EMISSION", "AnnualEmissions"))
    grouped = sub.groupby("YEAR")["AnnualEmissions"].sum()
    return {int(year): float(value) for year, value in grouped.items()}


def storage_series(
    scenario_frame: pd.DataFrame, region: str | None, storage: tuple[str, ...]
) -> dict[int, float]:
    """Total storage capacity by year for the profile's storage families."""
    required = {"TotalCapacityAnnual", "YEAR", "TECHNOLOGY"}
    if not required.issubset(scenario_frame.columns):
        return {}
    sub = scenario_frame[
        scenario_frame["TotalCapacityAnnual"].notna()
        & scenario_frame["YEAR"].notna()
        & scenario_frame["TECHNOLOGY"].notna()
    ]
    sub = sub[sub["TECHNOLOGY"].map(lambda tech: is_storage_tech(tech, storage)).astype(bool)]
    sub = _region_rows(sub, region)
    sub = _dedup(sub, ("YEAR", "TECHNOLOGY", "TotalCapacityAnnual"))
    grouped = sub.groupby("YEAR")["TotalCapacityAnnual"].sum()
    return {int(year): round(float(value), 4) for year, value in grouped.items()}


def vre_share_series(
    scenario_frame: pd.DataFrame, region: str | None, storage: tuple[str, ...]
) -> dict[int, float]:
    """VRE production share of valid generation output per year (0.0-1.0)."""
    required = {"ProductionByTechnologyAnnual", "YEAR", "TECHNOLOGY"}
    if not required.issubset(scenario_frame.columns):
        return {}
    sub = scenario_frame[
        scenario_frame["ProductionByTechnologyAnnual"].notna()
        & scenario_frame["YEAR"].notna()
        & scenario_frame["TECHNOLOGY"].notna()
    ]
    sub = sub[sub["TECHNOLOGY"].map(lambda tech: is_generation_tech(tech, storage)).astype(bool)]
    sub = _region_rows(sub, region)
    sub = _dedup(sub, ("YEAR", "TECHNOLOGY", "ProductionByTechnologyAnnual"))
    if sub.empty:
        return {}
    out: dict[int, float] = {}
    for year, group in sub.groupby("YEAR"):
        total = float(group["ProductionByTechnologyAnnual"].sum())
        vre = float(
            group.loc[
                group["TECHNOLOGY"].map(is_vre_tech).astype(bool), "ProductionByTechnologyAnnual"
            ].sum()
        )
        out[int(year)] = round(vre / total, 4) if total > 0 else 0.0
    return out


def cost_kpis(scenario_frame: pd.DataFrame) -> dict[str, float]:
    """System-wide cost KPIs for a scenario (never region-filtered)."""
    out: dict[str, float] = {}
    for key, column in (
        ("total_discounted", "TotalDiscountedCost"),
        ("capex", "CapitalInvestment"),
    ):
        if column not in scenario_frame.columns:
            out[key] = 0.0
            continue
        sub = scenario_frame[scenario_frame[column].notna()]
        sub = _dedup(
            sub, ("Scenario", "REGION", "TECHNOLOGY", "YEAR", "EMISSION", column)
        )
        out[key] = round(float(pd.to_numeric(sub[column], errors="coerce").sum()), 4)
    return out


def lid_diagnostic(
    scenario_frame: pd.DataFrame,
    region: str | None,
    storage: tuple[str, ...],
    region_labels: Mapping[str, str],
) -> dict[str, dict[str, object]]:
    """Per-technology NewCapacity vs TotalAnnualMaxCapacityInvestment (lid)."""
    newcap_column = "NewCapacity"
    lid_column = "TotalAnnualMaxCapacityInvestment"
    if {"YEAR", "TECHNOLOGY"} - set(scenario_frame.columns):
        return {}
    if (
        newcap_column not in scenario_frame.columns
        and lid_column not in scenario_frame.columns
    ):
        return {}
    sub = scenario_frame[
        scenario_frame["YEAR"].notna() & scenario_frame["TECHNOLOGY"].notna()
    ]
    sub = sub[sub["TECHNOLOGY"].map(lambda tech: is_generation_tech(tech, storage)).astype(bool)]
    sub = _region_rows(sub, region)
    sub = _dedup(sub, ("YEAR", "TECHNOLOGY", newcap_column, lid_column))
    out: dict[str, dict[str, object]] = {}
    for tech, group in sub.groupby("TECHNOLOGY"):
        family = tech_family(tech)
        if not family:
            continue
        newcap: dict[int, float] = {}
        lid: dict[int, float] = {}
        for _, row in group.iterrows():
            year = int(row["YEAR"])
            if newcap_column in group.columns and pd.notna(row.get(newcap_column)):
                newcap[year] = round(float(row[newcap_column]), 4)
            if lid_column in group.columns and pd.notna(row.get(lid_column)):
                lid[year] = round(float(row[lid_column]), 4)
        if all(value == 0 for value in newcap.values()) and all(
            value == 0 for value in lid.values()
        ):
            continue
        regions = technology_regions(tech)
        region_code = regions[0] if regions else ""
        region_label = region_labels.get(region_code, region_code)
        label = f"{family} — {region_label}" if region_label else family
        out[str(tech)] = {
            "label": label,
            "family": family,
            "newcap": newcap,
            "lid": lid,
        }
    return out


def _interconnector_block(
    scenario_frame: pd.DataFrame,
    region: str | None,
    technologies: Iterable[str],
) -> dict[str, dict[int, float]]:
    """{metric: {year: value}} over the profile's declared interconnectors."""
    declared = {str(tech) for tech in technologies}
    if not declared or {"YEAR", "TECHNOLOGY"} - set(scenario_frame.columns):
        return {}
    selected = scenario_frame[
        scenario_frame["TECHNOLOGY"].astype(str).isin(declared)
        & scenario_frame["YEAR"].notna()
    ]
    result: dict[str, dict[int, float]] = {}
    for metric in ("TotalCapacityAnnual", "ProductionByTechnologyAnnual"):
        if metric not in selected.columns:
            continue
        sub = selected[selected[metric].notna()]
        sub = _region_rows(sub, region)
        sub = _dedup(sub, ("YEAR", "TECHNOLOGY", metric))
        grouped = sub.groupby("YEAR")[metric].sum()
        series = {int(year): float(value) for year, value in grouped.items()}
        if series:
            result[metric] = series
    return result


def _region_label_map(country_regions: list[dict[str, str]]) -> dict[str, str]:
    labels = {"System": "System"}
    for entry in country_regions:
        labels[entry["region"]] = entry["label"]
    return labels


def _normalized_country_regions(
    metadata: Mapping[str, object],
) -> list[dict[str, str]]:
    raw = metadata.get("country_regions", [])
    if isinstance(raw, Mapping):
        items: list[object] = [
            {"region": str(region), "label": str(label)}
            for region, label in raw.items()
        ]
    elif isinstance(raw, list):
        items = raw
    else:
        raise ValueError("profile metadata country_regions must be a list or mapping")
    normalized: list[dict[str, str]] = []
    seen: set[str] = set()
    for item in items:
        if isinstance(item, Mapping):
            region = item.get("region")
            label = item.get("label", region)
        else:
            region = item
            label = item
        if not isinstance(region, str) or not region.strip():
            raise ValueError(f"invalid country region metadata: {item!r}")
        region = region.strip()
        if region in seen:
            raise ValueError(f"duplicate country region metadata: {region}")
        seen.add(region)
        normalized.append(
            {"region": region, "label": str(label) if label is not None else region}
        )
    return normalized


def _year_range(
    metadata: Mapping[str, object], observed_years: list[int]
) -> dict[str, int]:
    declared = metadata.get("year_range")
    if isinstance(declared, Mapping):
        try:
            start = int(declared.get("start"))  # type: ignore[arg-type]
            end = int(declared.get("end"))  # type: ignore[arg-type]
        except (TypeError, ValueError):
            pass
        else:
            if start <= end:
                return {"start": start, "end": end}
    if observed_years:
        return {"start": min(observed_years), "end": max(observed_years)}
    return dict(DEFAULT_YEAR_RANGE)


def aggregate_metrics(
    frame: pd.DataFrame,
    *,
    scenario: str,
    region: str | None,
    interconnectors: Iterable[Mapping[str, object] | str] = (),
    metadata: Mapping[str, object] | None = None,
) -> dict[str, object]:
    """All metric series/KPIs for one (scenario, region); missing data stays empty."""
    metadata = metadata or {}
    storage = storage_prefixes(metadata)
    region_labels = _region_label_map(_normalized_country_regions(metadata))
    declared = interconnector_metadata(interconnectors)
    selected = _scenario_frame(frame, scenario)
    return {
        "available": scenario_has_results(selected),
        "generation": _stacked_by_family(
            selected, "ProductionByTechnologyAnnual", region, storage
        ),
        "capacity": _stacked_by_family(
            selected, "TotalCapacityAnnual", region, storage
        ),
        "new_capacity": _stacked_by_family(selected, "NewCapacity", region, storage),
        "emissions": emissions_series(selected, region),
        "storage": storage_series(selected, region, storage),
        "vre_share": vre_share_series(selected, region, storage),
        "cost": cost_kpis(selected),
        "interconnectors": _interconnector_block(
            selected, region, [str(item["technology"]) for item in declared]
        ),
        "lid_diagnostic": lid_diagnostic(selected, region, storage, region_labels),
    }


def build_dashboard_data(
    snapshots: Iterable[tuple[str, Path]],
    *,
    profile_id: str,
    manifest: Path,
    workspace: Path,
    metadata: Mapping[str, object],
) -> dict[str, object]:
    """Build the ostram-profile-report-v1 payload for the given snapshots."""
    country_regions = _normalized_country_regions(metadata)
    regions = [entry["region"] for entry in country_regions]
    interconnectors = metadata.get("interconnectors", [])
    if not isinstance(interconnectors, list):
        raise ValueError("profile metadata interconnectors must be a list")
    declared = interconnector_metadata(interconnectors)
    scenario_hint = metadata.get("scenarios", [])
    data: dict[str, object] = {}
    snapshot_labels: list[str] = []
    scenario_ids: list[str] = []
    observed_years: list[int] = []
    for label, path in snapshots:
        label = str(label)
        if label in data:
            raise ValueError(f"duplicate snapshot label: {label}")
        frame = _load_frame(Path(path))
        if "YEAR" in frame.columns:
            years = frame["YEAR"].dropna()
            if not years.empty:
                observed_years.append(int(years.min()))
                observed_years.append(int(years.max()))
        scenarios = (
            sorted(frame["Scenario"].dropna().astype(str).unique())
            if "Scenario" in frame.columns
            else [str(item) for item in scenario_hint]
        )
        for scenario in scenarios:
            if scenario not in scenario_ids:
                scenario_ids.append(scenario)
        snapshot_labels.append(label)
        data[label] = {
            scenario: {
                region: aggregate_metrics(
                    frame,
                    scenario=scenario,
                    region=region,
                    interconnectors=interconnectors,
                    metadata=metadata,
                )
                for region in ["System", *regions]
            }
            for scenario in scenarios
        }
    title = str(metadata.get("title") or f"OSTRAM {profile_id} training report")
    return {
        "schema": "ostram-profile-report-v1",
        "generated_at_utc": datetime.now(timezone.utc).isoformat(),
        "profile_id": profile_id,
        "title": title,
        "manifest": str(Path(manifest).resolve()),
        "workspace": str(Path(workspace).resolve()),
        "country_regions": country_regions,
        "regions": [
            {"id": "System", "label": "System"},
            *[
                {"id": entry["region"], "label": entry["label"]}
                for entry in country_regions
            ],
        ],
        "scenarios": [
            {"id": scenario, "label": scenario_label(scenario)}
            for scenario in scenario_ids
        ],
        "families": [
            {"name": family, "color": tech_color(family)}
            for family in ORDERED_FAMILIES
        ],
        "interconnectors": declared,
        "effective_values": dict(metadata.get("effective_values", {})),
        "year_range": _year_range(metadata, observed_years),
        "snapshot_labels": snapshot_labels,
        "snapshots": data,
    }


_DASH_JS = r"""
(function(){
  "use strict";
  const D = JSON.parse(document.getElementById("ostram-profile-data").textContent);
  const LABELS = (D.snapshot_labels && D.snapshot_labels.length)
    ? D.snapshot_labels : Object.keys(D.snapshots || {});
  const SCENARIOS = D.scenarios || [];
  const REGIONS = (D.regions && D.regions.length)
    ? D.regions : [{id: "System", label: "System"}];
  const FAMILIES = D.families || [];
  const YR = D.year_range || {};
  const Y0 = Number.isFinite(YR.start) ? YR.start : 2023;
  const Y1 = (Number.isFinite(YR.end) && YR.end > Y0) ? YR.end : Y0 + 1;
  const $ = id => document.getElementById(id);
  const app = $("app");

  function esc(value){
    return String(value)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;")
      .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }

  if(!LABELS.length || !SCENARIOS.length){
    app.innerHTML = '<div class="panel"><div class="empty">' +
      'No captured result snapshots are available yet — run a scenario ' +
      'and capture it first.</div></div>';
    return;
  }

  function fillSelect(el, items, valueFn, labelFn, def){
    el.innerHTML = "";
    items.forEach(it => {
      const o = document.createElement("option");
      o.value = valueFn(it); o.textContent = labelFn(it);
      if(valueFn(it) === def) o.selected = true;
      el.appendChild(o);
    });
  }

  function defaultBefore(){
    const original = LABELS.find(l => String(l).toLowerCase() === "original");
    if(original !== undefined) return original;
    const baseline = LABELS.find(l => String(l).toLowerCase() === "baseline");
    return baseline !== undefined ? baseline : LABELS[0];
  }

  fillSelect($("sel-before"), LABELS, x => x, x => x, defaultBefore());
  fillSelect($("sel-after"), LABELS, x => x, x => x, LABELS[LABELS.length - 1]);
  fillSelect($("sel-scenario"), SCENARIOS, x => x.id, x => x.label, SCENARIOS[0].id);
  fillSelect($("sel-region"), REGIONS, x => x.id, x => x.label, "System");
  fillSelect($("sel-snapshot"), LABELS, x => x, x => x, LABELS[LABELS.length - 1]);
  fillSelect($("sel-scen-a"), SCENARIOS, x => x.id, x => x.label, SCENARIOS[0].id);
  fillSelect($("sel-scen-b"), SCENARIOS, x => x.id, x => x.label,
    SCENARIOS.length > 1 ? SCENARIOS[1].id : SCENARIOS[0].id);

  let currentMode = "snap";
  function setMode(m){
    currentMode = m;
    $("mode-snap").classList.toggle("active", m === "snap");
    $("mode-scen").classList.toggle("active", m === "scen");
    document.querySelectorAll(".ctrl-snap").forEach(el => {
      el.style.display = m === "snap" ? "" : "none";
    });
    document.querySelectorAll(".ctrl-scen").forEach(el => {
      el.style.display = m === "scen" ? "" : "none";
    });
    render();
  }
  $("mode-snap").addEventListener("click", () => setMode("snap"));
  $("mode-scen").addEventListener("click", () => setMode("scen"));

  ["sel-before", "sel-after", "sel-scenario", "sel-region",
   "sel-snapshot", "sel-scen-a", "sel-scen-b"].forEach(id => {
    $(id).addEventListener("change", render);
  });

  function metricsFor(snap, sc, region){
    const bySc = D.snapshots[snap];
    if(!bySc) return null;
    const byRegion = bySc[sc];
    if(!byRegion) return null;
    const m = byRegion[region];
    if(!m || !m.available) return null;
    return m;
  }

  // SVG helpers. Series are non-negative model outputs; the y-scale floor is 0.
  const W = 420, H = 260, PAD_L = 48, PAD_B = 28, PAD_T = 10, PAD_R = 10;
  function xScale(y){
    return PAD_L + (y - Y0) / (Y1 - Y0) * (W - PAD_L - PAD_R);
  }
  function yScale(v, vmax){
    return H - PAD_B - (vmax <= 0 ? 0 : v / vmax * (H - PAD_B - PAD_T));
  }
  function xTicks(){
    const out = [];
    for(let i = 0; i <= 3; i++){
      const y = Math.round(Y0 + (Y1 - Y0) * i / 3);
      if(out.indexOf(y) < 0) out.push(y);
    }
    return out;
  }
  function fmt(v){
    return Math.abs(v) >= 1000 ? (v / 1000).toFixed(1) + "k"
      : (Math.abs(v) >= 10 ? v.toFixed(0) : v.toFixed(1));
  }
  function fmtTip(v){ return String(Math.round(v * 10000) / 10000); }
  function axis(vmax, unit){
    let g = '<line x1="' + PAD_L + '" y1="' + (H - PAD_B) + '" x2="' + (W - PAD_R) +
      '" y2="' + (H - PAD_B) + '" stroke="#cbd5e1"/>' +
      '<line x1="' + PAD_L + '" y1="' + PAD_T + '" x2="' + PAD_L +
      '" y2="' + (H - PAD_B) + '" stroke="#cbd5e1"/>';
    xTicks().forEach(y => {
      g += '<text x="' + xScale(y) + '" y="' + (H - PAD_B + 16) +
        '" font-size="10" text-anchor="middle" fill="#64748b">' + y + '</text>';
    });
    for(let i = 0; i <= 2; i++){
      const v = vmax * i / 2;
      const yy = yScale(v, vmax);
      g += '<text x="' + (PAD_L - 6) + '" y="' + (yy + 3) +
        '" font-size="9" text-anchor="end" fill="#64748b">' + fmt(v) + '</text>';
      g += '<line x1="' + PAD_L + '" y1="' + yy + '" x2="' + (W - PAD_R) +
        '" y2="' + yy + '" stroke="#eef2f2"/>';
    }
    g += '<text x="' + PAD_L + '" y="' + (PAD_T - 1) +
      '" font-size="9" fill="#94a3b8">' + esc(unit) + '</text>';
    return g;
  }

  function years(obj){ return Object.keys(obj).map(Number).sort((a, b) => a - b); }

  function stackedArea(byYear, families, vmax){
    const ys = years(byYear);
    if(!ys.length) return '<div class="empty">No data</div>';
    let svg = '<svg viewBox="0 0 ' + W + ' ' + H + '">' + axis(vmax, "");
    const cum = {};
    ys.forEach(y => { cum[y] = 0; });
    families.forEach(f => {
      const top = [], bot = [];
      ys.forEach(y => {
        const val = (byYear[y] && byYear[y][f.name]) || 0;
        const b = cum[y]; const t = b + val;
        bot.push([y, b]); top.push([y, t]); cum[y] = t;
      });
      if(top.every((p, i) => p[1] === bot[i][1])) return;
      const pts = top.map(p => xScale(p[0]) + "," + yScale(p[1], vmax))
        .concat(bot.reverse().map(p => xScale(p[0]) + "," + yScale(p[1], vmax)))
        .join(" ");
      svg += '<polygon points="' + pts + '" fill="' + f.color +
        '" opacity="0.9"><title>' + esc(f.name) + '</title></polygon>';
    });
    svg += "</svg>";
    return svg;
  }

  function linePath(obj, vmax, color, dash){
    const ys = years(obj);
    if(!ys.length) return "";
    const d = ys.map((y, i) =>
      (i ? "L" : "M") + xScale(y) + "," + yScale(obj[y], vmax)).join(" ");
    let g = '<path d="' + d + '" fill="none" stroke="' + color +
      '" stroke-width="2"' + (dash ? ' stroke-dasharray="5,4"' : "") + '/>';
    ys.forEach(y => {
      g += '<circle cx="' + xScale(y) + '" cy="' + yScale(obj[y], vmax) +
        '" r="2.5" fill="' + color + '"><title>' + y + ": " +
        fmtTip(obj[y]) + '</title></circle>';
    });
    return g;
  }

  function lineChart(before, after, vmax, unit){
    if(!years(before).length && !years(after).length){
      return '<div class="empty">No data</div>';
    }
    return '<svg viewBox="0 0 ' + W + ' ' + H + '">' + axis(vmax, unit) +
      linePath(before, vmax, "#94a3b8", true) +
      linePath(after, vmax, "#0A595F", false) + "</svg>";
  }

  function maxOfStacks(){
    let m = 0;
    [...arguments].forEach(byYear => {
      if(!byYear) return;
      years(byYear).forEach(y => {
        const s = Object.values(byYear[y]).reduce((a, b) => a + b, 0);
        if(s > m) m = s;
      });
    });
    return m || 1;
  }
  function maxOfSeries(){
    let m = 0;
    [...arguments].forEach(o => {
      if(!o) return;
      Object.values(o).forEach(v => { if(v > m) m = v; });
    });
    return m || 1;
  }

  function legend(families, byYearA, byYearB){
    const present = new Set();
    [byYearA, byYearB].forEach(o => {
      if(!o) return;
      years(o).forEach(y => Object.keys(o[y]).forEach(f => present.add(f)));
    });
    return '<div class="legend">' + families.filter(f => present.has(f.name))
      .map(f => '<span><i style="background:' + f.color + '"></i>' +
        esc(f.name) + "</span>").join("") + "</div>";
  }

  function kpiCard(label, valAfter, valBefore, unit, lowerIsBetter){
    let deltaHtml = "";
    if(valBefore !== null && valBefore !== undefined && valBefore !== 0){
      const pct = (valAfter - valBefore) / Math.abs(valBefore) * 100;
      const cls = Math.abs(pct) < 0.05 ? "flat"
        : ((pct < 0) === lowerIsBetter ? "good" : "bad");
      const sign = pct > 0 ? "+" : "";
      deltaHtml = '<div class="k-delta ' + cls + '">' + sign +
        pct.toFixed(1) + "% vs Before</div>";
    }
    return '<div class="kpi"><div class="k-label">' + esc(label) + "</div>" +
      '<div class="k-value">' + fmt(valAfter) + (unit ? " " + esc(unit) : "") +
      "</div>" + deltaHtml + "</div>";
  }

  function pair(title, htmlBefore, htmlAfter, legendHtml){
    return '<div class="panel"><h2>' + title + '</h2><div class="pair">' +
      '<div class="col"><h3>Before</h3>' + htmlBefore + "</div>" +
      '<div class="col"><h3>After</h3>' + htmlAfter + "</div></div>" +
      (legendHtml || "") + "</div>";
  }

  function icSeries(m, metric){
    return (m && m.interconnectors && m.interconnectors[metric]) || null;
  }
  function lastPoint(o){
    const ys = years(o || {});
    if(!ys.length) return null;
    const y = ys[ys.length - 1];
    return {year: y, value: o[y]};
  }
  function scaleShare(o){
    if(!o) return {};
    const r = {};
    Object.keys(o).forEach(y => { r[y] = o[y] * 100; });
    return r;
  }

  function render(){
    const region = $("sel-region").value;
    let mB, mA;
    if(currentMode === "snap"){
      const before = $("sel-before").value, after = $("sel-after").value,
            sc = $("sel-scenario").value;
      mB = metricsFor(before, sc, region);
      mA = metricsFor(after, sc, region);
    } else {
      const snap = $("sel-snapshot").value,
            scA = $("sel-scen-a").value, scB = $("sel-scen-b").value;
      mB = metricsFor(snap, scA, region);
      mA = metricsFor(snap, scB, region);
    }
    app.innerHTML = "";

    if(!mA && !mB){
      app.innerHTML = '<div class="panel"><div class="empty">No results for ' +
        'this scenario in the selected snapshots — run it first.</div></div>';
      return;
    }

    // KPI cards
    const cA = (mA && mA.cost) || {}, cB = (mB && mB.cost) || {};
    const lpA = lastPoint(mA && mA.vre_share);
    const vreYear = lpA ? lpA.year : (lastPoint(mB && mB.vre_share) || {}).year;
    const vreA = lpA ? lpA.value : (mA && mA.vre_share ? mA.vre_share[vreYear] : null);
    const vreB = (mB && mB.vre_share && vreYear !== undefined)
      ? mB.vre_share[vreYear] : null;
    let kpis = '<div class="kpis">';
    kpis += kpiCard("System cost (NPV)", cA.total_discounted || 0,
      cB.total_discounted, "", true);
    kpis += kpiCard("Capital investment", cA.capex || 0, cB.capex, "", true);
    kpis += kpiCard("VRE share" + (vreYear !== undefined ? " @" + vreYear : ""),
      vreA !== null && vreA !== undefined ? vreA * 100 : 0,
      vreB !== null && vreB !== undefined ? vreB * 100 : null, "%", false);
    kpis += "</div>";
    app.innerHTML += kpis;

    // Generation mix
    const gB = mB && mB.generation, gA = mA && mA.generation;
    const gMax = maxOfStacks(gB, gA);
    app.innerHTML += pair("Generation mix by technology (PJ)",
      gB ? stackedArea(gB, FAMILIES, gMax) : '<div class="empty">No Before data</div>',
      gA ? stackedArea(gA, FAMILIES, gMax) : '<div class="empty">No After data</div>',
      legend(FAMILIES, gA, gB));

    // Capacity mix
    const kB = mB && mB.capacity, kA = mA && mA.capacity;
    const kMax = maxOfStacks(kB, kA);
    app.innerHTML += pair("Installed capacity by technology (GW)",
      kB ? stackedArea(kB, FAMILIES, kMax) : '<div class="empty">No Before data</div>',
      kA ? stackedArea(kA, FAMILIES, kMax) : '<div class="empty">No After data</div>',
      legend(FAMILIES, kA, kB));

    // New capacity (annual investment)
    const ncB = mB && mB.new_capacity, ncA = mA && mA.new_capacity;
    if((ncB && years(ncB).length) || (ncA && years(ncA).length)){
      const ncMax = maxOfStacks(ncB, ncA);
      app.innerHTML += pair("New capacity investment by technology (GW per year)",
        ncB ? stackedArea(ncB, FAMILIES, ncMax)
          : '<div class="empty">No Before data</div>',
        ncA ? stackedArea(ncA, FAMILIES, ncMax)
          : '<div class="empty">No After data</div>',
        legend(FAMILIES, ncA, ncB));
    }

    // CO2, storage as overlaid line charts
    const eMax = maxOfSeries(mB && mB.emissions, mA && mA.emissions);
    app.innerHTML += '<div class="panel"><h2>CO₂ emissions ' +
      "(dashed = Before, solid = After)</h2>" +
      lineChart((mB && mB.emissions) || {}, (mA && mA.emissions) || {}, eMax, "") +
      "</div>";
    const sMax = maxOfSeries(mB && mB.storage, mA && mA.storage);
    app.innerHTML += '<div class="panel"><h2>Storage capacity (GW) — ' +
      "dashed = Before, solid = After</h2>" +
      lineChart((mB && mB.storage) || {}, (mA && mA.storage) || {}, sMax, "GW") +
      "</div>";

    // Interconnector capacity and flow
    const icB = icSeries(mB, "TotalCapacityAnnual");
    const icA = icSeries(mA, "TotalCapacityAnnual");
    if(icB || icA){
      const icMax = maxOfSeries(icB, icA);
      const seed = D.effective_values
        ? D.effective_values.interconnector_capacity_gw : null;
      const seedNote = (seed === null || seed === undefined) ? ""
        : '<p class="note">Declared effective interconnector seed capacity: ' +
          esc(seed) + " GW.</p>";
      app.innerHTML += '<div class="panel"><h2>Interconnector capacity (GW) ' +
        "— dashed = Before, solid = After</h2>" +
        lineChart(icB || {}, icA || {}, icMax, "GW") + seedNote + "</div>";
    }
    const ipB = icSeries(mB, "ProductionByTechnologyAnnual");
    const ipA = icSeries(mA, "ProductionByTechnologyAnnual");
    if(ipB || ipA){
      const ipMax = maxOfSeries(ipB, ipA);
      app.innerHTML += '<div class="panel"><h2>Interconnector trade flow (PJ) ' +
        "— dashed = Before, solid = After</h2>" +
        lineChart(ipB || {}, ipA || {}, ipMax, "PJ") + "</div>";
    }

    // VRE share
    app.innerHTML += '<div class="panel"><h2>VRE share of generation — ' +
      "dashed = Before, solid = After</h2>" +
      lineChart(scaleShare(mB && mB.vre_share), scaleShare(mA && mA.vre_share),
        100, "%") + "</div>";

    // Lid diagnostic
    const lidA = mA && mA.lid_diagnostic, lidB = mB && mB.lid_diagnostic;
    if(lidA || lidB){
      const allTechs = {};
      [lidB, lidA].forEach(d => {
        if(d) Object.keys(d).forEach(t => { allTechs[t] = d[t].label; });
      });
      const techList = Object.entries(allTechs)
        .sort((a, b) => a[1].localeCompare(b[1]));
      if(techList.length){
        let h = '<div class="panel"><h2>Lid diagnostic — NewCapacity vs ' +
          "MaxCapacityInvestment (GW)</h2>" +
          '<p class="note">Pick a technology. When the solid line touches the ' +
          "dashed line the technology is <strong>lid-bound</strong> (the lid " +
          "decides). A gap means <strong>economics-bound</strong> (the " +
          "optimiser chose less than the lid allows).</p>" +
          '<div style="margin:12px 0"><label class="lid-label">Technology' +
          '<select id="sel-lid-tech">';
        techList.forEach(([code, label]) => {
          h += '<option value="' + esc(code) + '">' + esc(label) + "</option>";
        });
        h += "</select></label></div>" +
          '<div class="pair"><div class="col"><h3>Before</h3>' +
          '<div id="lid-before"></div></div>' +
          '<div class="col"><h3>After</h3><div id="lid-after"></div></div>' +
          "</div></div>";
        app.innerHTML += h;

        const renderLid = function(){
          const t = $("sel-lid-tech").value;
          const dB = lidB && lidB[t], dA = lidA && lidA[t];
          function draw(d){
            if(!d) return '<div class="empty">No data for this technology</div>';
            const nc = d.newcap || {}, lid = d.lid || {};
            const vm = Math.max(maxOfSeries(nc), maxOfSeries(lid)) || 1;
            if(!years(nc).length && !years(lid).length){
              return '<div class="empty">No data</div>';
            }
            return '<svg viewBox="0 0 ' + W + " " + H + '">' + axis(vm, "GW") +
              linePath(lid, vm, "#F59E0B", true) +
              linePath(nc, vm, "#0A595F", false) + "</svg>" +
              '<div class="legend"><span><i style="background:#0A595F"></i>' +
              "NewCapacity</span>" +
              '<span><i style="background:#F59E0B"></i>' +
              "MaxCapInvestment (lid)</span></div>";
          }
          $("lid-before").innerHTML = draw(dB);
          $("lid-after").innerHTML = draw(dA);
        };
        $("sel-lid-tech").addEventListener("change", renderLid);
        renderLid();
      }
    }
  }

  render();
})();
"""


_HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>__TITLE__ — Results Dashboard</title>
<style>
  *,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
  :root {
    --teal-dark:#00414D; --teal:#0A595F; --teal-mid:#23978E;
    --emerald:#10B981; --amber:#F59E0B; --slate-900:#334155;
    --slate-600:#475569; --slate-400:#94A3B8; --sage-light:#C5D8D8;
    --mint-light:#D0EDEB; --cloud:#E0F0EF; --ice:#F0F5F5;
    --white:#FFFFFF; --red:#EF4444;
  }
  body{font-family:'Inter',-apple-system,BlinkMacSystemFont,sans-serif;
    background:var(--ice);color:var(--slate-900);line-height:1.6;font-size:15px}

  /* Hero */
  .hero { background:linear-gradient(135deg,var(--teal-dark) 0%,var(--teal) 50%,var(--teal-mid) 100%); color:var(--white); padding:48px 40px 40px; text-align:center; position:relative; overflow:hidden; }
  .hero::before { content:''; position:absolute; top:-60px; right:-60px; width:250px; height:250px; border-radius:50%; background:rgba(16,185,129,.12); }
  .hero::after  { content:''; position:absolute; bottom:-40px; left:-40px; width:180px; height:180px; border-radius:50%; background:rgba(245,158,11,.1); }
  .hero h1 { font-size:2.2rem; font-weight:800; letter-spacing:-.5px; margin-bottom:10px; position:relative; }
  .hero .subtitle { font-size:1.05rem; opacity:.88; max-width:680px; margin:0 auto; position:relative; }
  .hero .badge { display:inline-block; background:rgba(255,255,255,.15); border:1px solid rgba(255,255,255,.25); border-radius:20px; padding:4px 16px; font-size:.82rem; font-weight:500; margin-bottom:18px; position:relative; }

  /* Controls */
  .controls{display:flex;flex-wrap:wrap;gap:16px;padding:18px 32px;align-items:flex-end;
    background:var(--white);border-bottom:1px solid var(--sage-light);position:sticky;top:0;z-index:5}
  .controls label{display:flex;flex-direction:column;font-size:.72rem;
    font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--teal)}
  .controls select{margin-top:4px;padding:7px 10px;border:1px solid var(--sage-light);
    border-radius:8px;font-size:.9rem;background:var(--ice);color:var(--slate-900);min-width:150px}
  .compare-mode{display:flex;gap:0;border:1px solid var(--sage-light);border-radius:8px;overflow:hidden;align-self:flex-end;margin-bottom:2px}
  .mode-btn{border:none;padding:7px 14px;font-size:.78rem;font-weight:600;font-family:inherit;
    cursor:pointer;background:var(--ice);color:var(--slate-600);transition:all .15s}
  .mode-btn.active{background:var(--teal);color:var(--white)}
  .mode-btn:hover:not(.active){background:var(--cloud)}

  main{max-width:1180px;margin:0 auto;padding:28px 32px}
  .kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:16px;margin-bottom:24px}
  .kpi{background:var(--white);border:1px solid var(--sage-light);border-radius:12px;padding:16px 18px;box-shadow:0 2px 10px rgba(0,65,77,.04)}
  .kpi .k-label{font-size:.72rem;font-weight:700;text-transform:uppercase;color:var(--teal);letter-spacing:.5px}
  .kpi .k-value{font-size:1.6rem;font-weight:800;color:var(--slate-900);margin-top:4px}
  .kpi .k-delta{font-size:.85rem;font-weight:600;margin-top:2px}
  .k-delta.good{color:var(--emerald)} .k-delta.bad{color:var(--red)} .k-delta.flat{color:var(--slate-600)}
  .panel{background:var(--white);border:1px solid var(--sage-light);border-radius:12px;padding:24px;margin-bottom:24px;box-shadow:0 2px 10px rgba(0,65,77,.04)}
  .panel h2{font-size:1.1rem;font-weight:800;color:var(--teal-dark);margin-bottom:14px}
  .pair{display:grid;grid-template-columns:1fr 1fr;gap:18px}
  .pair .col h3{font-size:.85rem;font-weight:700;color:var(--slate-600);text-align:center;margin-bottom:6px}
  .empty{padding:40px;text-align:center;color:var(--slate-600);font-style:italic}
  .legend{display:flex;flex-wrap:wrap;gap:10px;margin-top:10px;font-size:.75rem}
  .legend span{display:inline-flex;align-items:center;gap:5px}
  .legend i{width:12px;height:12px;border-radius:3px;display:inline-block}
  svg{width:100%;height:auto;display:block}
  .note{font-size:.8rem;color:var(--slate-600);margin-top:8px}
  .lid-label{font-size:.72rem;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--teal)}
  .lid-label select{margin-left:8px;padding:7px 10px;border:1px solid var(--sage-light);border-radius:8px;font-size:.9rem;background:var(--ice);color:var(--slate-900);min-width:260px}
  @media(max-width:760px){.pair{grid-template-columns:1fr}}

  /* Footer */
  footer { background:var(--teal-dark); color:rgba(255,255,255,.75); padding:40px 24px 30px; margin-top:10px; }
  .footer-inner { max-width:1180px; margin:0 auto; }
  .footer-brand { display:flex; align-items:flex-start; gap:20px; margin-bottom:24px; padding-bottom:24px; border-bottom:1px solid rgba(255,255,255,.15); flex-wrap:wrap; }
  .footer-logo-box { background:var(--white); border-radius:10px; padding:10px 16px; display:flex; align-items:center; justify-content:center; flex-shrink:0; }
  .footer-logo-svg { width:180px; height:auto; }
  .footer-tagline { flex:1; min-width:220px; }
  .footer-tagline strong { display:block; color:var(--white); font-size:1rem; font-weight:700; margin-bottom:4px; }
  .footer-tagline p { font-size:.88rem; line-height:1.6; margin:0; color:rgba(255,255,255,.7); }
  .footer-bottom { font-size:.8rem; color:rgba(255,255,255,.45); text-align:center; padding-top:20px; border-top:1px solid rgba(255,255,255,.1); }
  .footer-meta { font-size:.75rem; color:rgba(255,255,255,.4); text-align:center; padding-top:10px; overflow-wrap:anywhere; }

  @media (max-width:700px) {
    .hero h1 { font-size:1.7rem; }
  }
</style>
</head>
<body>

<div class="hero">
  <div class="badge">__TITLE__</div>
  <h1>Results Dashboard</h1>
  <p class="subtitle">Compare a scenario's results <strong>Before</strong> vs <strong>After</strong> your edits &mdash; or switch to <strong>Scenarios</strong> mode to compare two scenarios side by side. Pick your region below.</p>
</div>

<div class="controls">
  <div class="compare-mode">
    <button id="mode-snap" class="mode-btn active">Snapshots</button>
    <button id="mode-scen" class="mode-btn">Scenarios</button>
  </div>
  <label class="ctrl-snap">Before<select id="sel-before"></select></label>
  <label class="ctrl-snap">After<select id="sel-after"></select></label>
  <label class="ctrl-snap">Scenario<select id="sel-scenario"></select></label>
  <label class="ctrl-scen" style="display:none">Snapshot<select id="sel-snapshot"></select></label>
  <label class="ctrl-scen" style="display:none">Before scenario<select id="sel-scen-a"></select></label>
  <label class="ctrl-scen" style="display:none">After scenario<select id="sel-scen-b"></select></label>
  <label>Region<select id="sel-region"></select></label>
</div>
<main id="app"></main>
<script id="ostram-profile-data" type="application/json">__PAYLOAD__</script>
<script>
__JS__
</script>

<footer>
  <div class="footer-inner">
    <div class="footer-brand">
      <div class="footer-logo-box">
<svg class="footer-logo-svg" viewBox="55 -5 560 310" aria-label="Climate Lead Group">
<path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M8723.98 26126.6C8774.45 26126.6 8820.88 26136.3 8861.99 26155.4 8903.26 26174.6 8938.21 26202.2 8965.86 26237.4L8975.79 26250 8909.67 26300.1 8899.99 26288.9C8853.81 26235.1 8795.67 26207.9 8727.2 26207.9 8664.84 26207.9 8612.57 26228 8571.86 26267.7 8531.27 26307.3 8510.69 26357.5 8510.69 26416.9 8510.69 26456.3 8519.89 26492.1 8538.05 26523.5 8556.17 26554.8 8582.14 26580.2 8615.23 26598.8 8648.4 26617.5 8684.99 26627 8723.98 26627 8759.79 26627 8793.02 26619.7 8822.76 26605.5 8852.4 26591.3 8878.2 26570 8899.45 26542.3L8909.07 26529.8 8975.22 26580.6 8966.25 26593C8939.59 26630 8904.91 26659.1 8863.18 26679.6 8821.64 26700 8774.56 26710.4 8723.27 26710.4 8638.12 26710.4 8566.18 26682.3 8509.43 26627.1 8452.54 26571.7 8423.69 26502.6 8423.69 26421.5 8423.69 26344.8 8447.47 26278.3 8494.39 26223.9 8550.41 26159.3 8627.65 26126.6 8723.98 26126.6" fill="#00414c"/><path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M9108.97 26710.4H9027.34V26139.8H9108.97V26710.4" fill="#00414c"/><path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M9173.87 26139.8H9255.5V26560.3H9173.87V26139.8" fill="#00414c"/><path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M9214.51 26606.2C9230.19 26606.2 9243.78 26611.9 9254.88 26623 9265.97 26634.2 9271.58 26647.7 9271.58 26663.1 9271.58 26678.8 9265.97 26692.4 9254.88 26703.6 9243.79 26714.7 9230.2 26720.4 9214.51 26720.4 9199.06 26720.4 9185.58 26714.7 9174.49 26703.6 9163.4 26692.4 9157.79 26678.8 9157.79 26663.1 9157.79 26647.6 9163.41 26634.1 9174.49 26623 9185.59 26611.9 9199.06 26606.2 9214.51 26606.2" fill="#00414c"/><path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M9782.9 26490.1C9799.58 26490.1 9813.95 26486.2 9825.64 26478.4 9837.13 26470.8 9845.35 26460.6 9850.05 26448.3 9853.65 26438.8 9857.92 26417.2 9857.92 26367.9V26139.8H9940.64V26367.9C9940.64 26418.3 9935.2 26457.1 9924.48 26483.3 9913.32 26510.5 9895.91 26532 9872.76 26547.3 9849.68 26562.5 9822.81 26570.3 9792.91 26570.3 9761.23 26570.3 9731.27 26561.4 9703.84 26544 9684.64 26531.8 9667.52 26515.5 9652.83 26495.3 9647.08 26506.8 9640.74 26516.5 9633.9 26524.3 9621.68 26538.3 9605.84 26549.6 9586.84 26557.9 9567.97 26566.1 9547.78 26570.3 9526.84 26570.3 9496.16 26570.3 9466.54 26561.8 9438.7 26545 9428.74 26538.8 9418.86 26530.8 9409.18 26521.1V26560.3H9327.55V26139.8H9409.18V26320.4C9409.18 26366.7 9413.66 26401.1 9422.48 26422.7 9430.91 26443.3 9444.04 26459.9 9461.53 26472 9478.88 26484 9497.68 26490.1 9517.41 26490.1 9534.05 26490.1 9548.44 26485.9 9560.18 26477.7 9572.03 26469.5 9580.1 26458.8 9584.84 26445.1 9588.46 26434.8 9592.74 26410.7 9592.74 26354.7V26139.8H9674.38V26307.9C9674.38 26361.4 9678.7 26399.7 9687.19 26421.5 9695.27 26442.3 9708.19 26459.1 9725.61 26471.6 9742.87 26483.8 9762.15 26490.1 9782.9 26490.1" fill="#00414c"/><path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M10079.3 26349.5C10079.3 26374.2 10085.6 26398.1 10098.1 26420.3 10110.6 26442.4 10127.6 26459.8 10148.6 26472.1 10169.6 26484.2 10192.8 26490.4 10217.7 26490.4 10256.4 26490.4 10289.4 26476.8 10315.9 26449.8 10342.4 26422.8 10355.8 26388.6 10355.8 26348.1 10355.8 26321.5 10349.8 26297.3 10338 26276.2 10326.3 26255.3 10309.3 26238.5 10287.5 26226.2 10243.2 26201.3 10191.5 26201.3 10149.2 26226.2 10128 26238.7 10110.9 26256.3 10098.3 26278.7 10085.7 26301.2 10079.3 26325 10079.3 26349.5ZM10212.2 26129.8C10245.5 26129.8 10276.6 26136.9 10304.8 26150.8 10321.6 26159.1 10337.6 26169.9 10352.6 26183.2V26139.8H10433.5V26560.3H10352.6V26515.4C10338.7 26528.7 10323.7 26539.7 10307.6 26548.1 10279.6 26562.8 10248 26570.3 10213.6 26570.3 10154.2 26570.3 10102.7 26548.6 10060.5 26505.9 10018.3 26463.2 9996.96 26411.1 9996.96 26351.1 9996.96 26289.9 10018.1 26237.2 10059.9 26194.4 10101.8 26151.6 10153 26129.8 10212.2 26129.8" fill="#00414c"/><path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M10639.7 26139.8V26485.4H10719.1V26560.3H10639.7V26704.7H10558.1V26560.3H10489.8V26485.4H10558.1V26139.8H10639.7" fill="#00414c"/><path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M11084.5 26405.2H10837.7C10846.6 26427.9 10858 26445.7 10871.8 26458 10896.2 26479.8 10925.7 26490.8 10959.7 26490.8 10980.4 26490.8 11000.4 26486.4 11019.2 26477.6 11037.9 26468.9 11053 26457.4 11064.2 26443.5 11072.2 26433.4 11079 26420.5 11084.5 26405.2ZM11092.1 26292.8C11077.8 26268.9 11064.6 26251.6 11052.7 26241.4 11041.1 26231.3 11026.7 26223 11009.9 26216.9 10993 26210.7 10975.4 26207.5 10957.5 26207.5 10920.7 26207.5 10890.5 26220.1 10865.5 26246 10843.9 26268.3 10831.2 26296.7 10827.6 26330.3H11176.7L11176.5 26346.3C11176 26402.8 11160.4 26450.7 11130.2 26488.8 11087.6 26542.9 11030 26570.3 10959 26570.3 10889.8 26570.3 10833.7 26543.5 10792.4 26490.6 10760 26449.4 10743.6 26401.7 10743.6 26349 10743.6 26292.9 10762.9 26242 10801.1 26197.7 10839.8 26152.6 10893.9 26129.8 10961.8 26129.8 10992.3 26129.8 11020 26134.6 11044.4 26143.9 11068.8 26153.3 11091.1 26167.1 11110.6 26185.1 11130.1 26203 11147.3 26226.6 11161.9 26255.3L11169 26269.2 11099.8 26305.6 11092.1 26292.8" fill="#00414c"/><path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M8475.4 25977.6H8423.26V25618.5H8606.63V25669.4H8475.4V25977.6" fill="#00414c"/><path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M8683.08 25788.1C8688.94 25803.9 8696.67 25816.1 8706.14 25824.6 8722.37 25839.1 8741.29 25846.1 8763.95 25846.1 8777.55 25846.1 8790.72 25843.2 8803.08 25837.5 8815.36 25831.8 8825.32 25824.2 8832.69 25815 8838.28 25808 8842.96 25799 8846.63 25788.1H8683.08ZM8903.4 25751.1C8903.06 25787.5 8893.04 25818.4 8873.61 25842.9 8846.25 25877.7 8809.2 25895.3 8763.48 25895.3 8718.99 25895.3 8682.94 25878 8656.36 25844.1 8635.53 25817.5 8624.97 25786.8 8624.97 25752.9 8624.97 25716.8 8637.42 25684.1 8661.97 25655.6 8686.85 25626.6 8721.63 25612 8765.34 25612 8784.93 25612 8802.81 25615 8818.47 25621 8834.16 25627.1 8848.48 25636 8861.05 25647.5 8873.54 25659 8884.65 25674.2 8894.06 25692.7L8898.11 25700.7 8855.32 25723.1 8850.93 25715.8C8841.56 25700.3 8832.88 25689 8825.12 25682.2 8817.46 25675.6 8807.98 25670.1 8796.94 25666.1 8785.85 25662 8774.27 25659.9 8762.56 25659.9 8738.27 25659.9 8718.45 25668.2 8701.96 25685.3 8687.26 25700.5 8679.03 25719.1 8676.92 25742H8903.48L8903.4 25751.1" fill="#00414c"/><path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M9141.02 25819.2C9158.68 25801.2 9167.27 25779.3 9167.27 25752.3 9167.27 25734.9 9163.33 25719 9155.57 25705.1 9147.86 25691.4 9136.68 25680.3 9122.32 25672.2 9093.55 25656 9059.29 25655.8 9031.38 25672.2 9017.47 25680.4 9006.23 25692 8997.97 25706.7 8989.67 25721.5 8985.47 25737.2 8985.47 25753.3 8985.47 25769.5 8989.64 25785.2 8997.85 25799.8 9006.05 25814.3 9017.2 25825.8 9030.99 25833.8 9044.83 25841.8 9060.13 25845.9 9076.48 25845.9 9102.29 25845.9 9123.4 25837.2 9141.02 25819.2ZM9165.19 25858C9155.72 25867.5 9145.37 25875.2 9134.23 25881 9116.23 25890.5 9095.9 25895.3 9073.82 25895.3 9035.62 25895.3 9002.5 25881.3 8975.35 25853.9 8948.24 25826.4 8934.49 25792.9 8934.49 25754.3 8934.49 25714.9 8948.11 25681 8974.96 25653.5 9001.88 25626 9034.83 25612 9072.89 25612 9094.29 25612 9114.31 25616.5 9132.41 25625.4 9144.05 25631.2 9155.01 25638.8 9165.19 25648.1V25618.5H9215.23V25888.8H9165.19V25858" fill="#00414c"/><path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M9465.83 25819.2C9483.52 25801.2 9492.1 25779.3 9492.1 25752.3 9492.1 25734.9 9488.15 25719 9480.38 25705.1 9472.69 25691.4 9461.5 25680.3 9447.13 25672.2 9418.39 25656 9384.13 25655.8 9356.21 25672.2 9342.31 25680.4 9331.07 25692 9322.8 25706.7 9314.5 25721.5 9310.29 25737.2 9310.29 25753.3 9310.29 25769.5 9314.46 25785.2 9322.69 25799.8 9330.87 25814.3 9342.02 25825.8 9355.82 25833.8 9369.65 25841.8 9384.94 25845.9 9401.3 25845.9 9427.12 25845.9 9448.22 25837.2 9465.83 25819.2ZM9490.01 25858C9480.54 25867.5 9470.19 25875.2 9459.06 25881 9441.04 25890.5 9420.72 25895.3 9398.64 25895.3 9360.44 25895.3 9327.32 25881.3 9300.17 25853.9 9273.06 25826.4 9259.32 25792.9 9259.32 25754.3 9259.32 25714.9 9272.94 25681 9299.79 25653.5 9326.7 25626 9359.64 25612 9397.71 25612 9419.11 25612 9439.14 25616.5 9457.23 25625.4 9468.87 25631.2 9479.83 25638.8 9490.01 25648.1V25986.1H9540.06V25618.5H9490.01V25858" fill="#00414c"/><path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M9934.47 25749H10041.4C10035.8 25724.3 10023.3 25704.3 10003.4 25688.1 9980.9 25669.7 9952.86 25660.4 9920.09 25660.4 9892.83 25660.4 9867.44 25666.7 9844.64 25679.1 9821.92 25691.5 9803.88 25708.7 9791.04 25730.3 9778.19 25751.9 9771.67 25774.9 9771.67 25798.8 9771.67 25822.1 9777.96 25844.5 9790.37 25865.5 9802.77 25886.5 9820.5 25903.4 9843.11 25915.6 9884.14 25937.9 9929.77 25941 9977.86 25920.5 9999.37 25911.3 10019.5 25897.6 10037.6 25879.8L10043.8 25873.7 10083.4 25911.3 10076.5 25917.8C10051.5 25941.5 10025.5 25959 9999.07 25969.8 9972.63 25980.7 9944.01 25986.1 9914.06 25986.1 9876.98 25986.1 9843.2 25977.8 9813.7 25961.5 9784.1 25945.1 9760.49 25922 9743.49 25892.9 9726.53 25863.8 9717.93 25831.9 9717.93 25798.1 9717.93 25752.7 9732.93 25712.1 9762.52 25677.5 9800.92 25632.6 9853.62 25609.9 9919.15 25609.9 9973.33 25609.9 10017.3 25626.4 10049.9 25659 10082.5 25691.6 10099.2 25736 10099.5 25790.9L10099.6 25800H9934.47V25749" fill="#219a90"/><path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M10269.7 25885.5C10243.2 25900.4 10219.3 25897.4 10197.3 25883 10191.4 25879.1 10185.6 25874.3 10180 25868.5V25888.8H10128.8V25618.5H10180V25712.8C10180 25755.5 10181.9 25784.1 10185.7 25797.6 10190.4 25814.5 10197.2 25826.9 10205.9 25834.6 10214.3 25842.1 10222.6 25845.7 10231.2 25845.7 10233.5 25845.7 10237.8 25845.1 10244.8 25842L10252 25838.9 10277.9 25880.9 10269.7 25885.5" fill="#219a90"/><path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M10501.8 25752.2C10501.8 25735.7 10497.7 25720 10489.7 25705.6 10481.7 25691.4 10470.8 25680.3 10457.3 25672.5 10430.1 25657 10394.2 25657.1 10367.1 25672.6 10353.6 25680.3 10342.6 25691.4 10334.7 25705.6 10326.6 25720 10322.6 25735.6 10322.6 25752.2 10322.6 25778.3 10331.2 25799.9 10349.1 25818.3 10366.9 25836.6 10387.5 25845.4 10412.2 25845.4 10436.7 25845.4 10457.3 25836.6 10475.2 25818.3 10493.1 25799.9 10501.8 25778.3 10501.8 25752.2ZM10412.3 25895.3C10370.8 25895.3 10336 25880 10308.8 25849.9 10284.1 25822.6 10271.6 25790 10271.6 25752.9 10271.6 25715.5 10284.8 25682.3 10310.8 25654.4 10337 25626.2 10371.1 25612 10412.3 25612 10453.3 25612 10487.3 25626.2 10513.5 25654.4 10539.6 25682.3 10552.8 25715.5 10552.8 25752.9 10552.8 25790.2 10540.3 25822.9 10515.6 25850.1 10488.4 25880.1 10453.6 25895.3 10412.3 25895.3" fill="#219a90"/><path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M10778.2 25762C10778.2 25724.1 10775.8 25708.9 10773.8 25702.9 10769.6 25690.5 10762.1 25680.7 10751.1 25673.1 10739.9 25665.3 10726.5 25661.6 10710.1 25661.6 10693.9 25661.6 10679.9 25665.6 10668.4 25673.4 10657 25681.3 10649.5 25691.6 10645.4 25705.1 10642.6 25715.1 10641.1 25734.2 10641.1 25762V25888.8H10590.6V25758.3C10590.6 25720.8 10594.9 25693.2 10603.6 25673.9 10612.4 25654.5 10626 25639.1 10644 25628.2 10661.8 25617.4 10684.1 25612 10710.2 25612 10736.4 25612 10758.6 25617.4 10776.3 25628.3 10794 25639.1 10807.5 25654.4 10816.4 25673.6 10825.1 25692.5 10829.4 25719.5 10829.4 25755.9V25888.8H10778.2V25762" fill="#219a90"/><path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M11113.7 25753.3C11113.7 25737.2 11109.5 25721.5 11101.2 25706.7 11092.9 25692 11081.7 25680.4 11067.8 25672.2 11039.5 25655.6 11006 25656 10977.1 25672.2 10962.7 25680.3 10951.5 25691.4 10943.8 25705.1 10936 25719 10932.1 25734.9 10932.1 25752.3 10932.1 25779.3 10940.7 25801.2 10958.3 25819.2 10975.9 25837.2 10997 25845.9 11022.7 25845.9 11039.1 25845.9 11054.4 25841.8 11068.2 25833.8 11082 25825.8 11093.1 25814.3 11101.3 25799.8 11109.5 25785.2 11113.7 25769.5 11113.7 25753.3ZM11025.4 25895.3C11003.5 25895.3 10983.3 25890.5 10965.4 25881 10954.3 25875.1 10943.9 25867.4 10934.4 25857.9V25888.8H10883.9V25526.2H10934.4V25648.1C10944.5 25638.8 10955.4 25631.2 10967 25625.4 10985 25616.5 11005 25612 11026.4 25612 11064.4 25612 11097.3 25626 11124.2 25653.5 11151 25681 11164.6 25714.9 11164.6 25754.3 11164.6 25792.9 11150.9 25826.4 11123.8 25853.9 11096.7 25881.3 11063.6 25895.3 11025.4 25895.3" fill="#219a90"/><path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M7378.75 27268.9C6927.57 27268.9 6560.52 26901.9 6560.52 26450.7 6560.52 26188.6 6687.55 25940.4 6900.34 25786.8 6915.47 25775.9 6936.67 25779.3 6947.6 25794.5 6958.52 25809.6 6955.09 25830.8 6939.97 25841.7 6744.76 25982.6 6628.22 26210.3 6628.22 26450.7 6628.22 26864.6 6964.91 27201.2 7378.75 27201.2 7792.6 27201.2 8129.29 26864.6 8129.29 26450.7 8129.29 26208.3 8011.16 25979.4 7813.29 25838.7 7805.92 25833.4 7801.04 25825.6 7799.53 25816.7 7798.03 25807.8 7800.08 25798.8 7805.33 25791.5 7811.67 25782.6 7821.99 25777.3 7832.94 25777.3 7839.99 25777.3 7846.77 25779.4 7852.54 25783.5 8068.22 25937 8196.99 26186.4 8196.99 26450.7 8196.99 26901.9 7829.92 27268.9 7378.75 27268.9" fill="#00414c"/><path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M7867.2 25594.4H7475.95V26454.2C7475.95 26472.9 7460.76 26488.1 7442.1 26488.1 7423.43 26488.1 7408.25 26472.9 7408.25 26454.2V25560.5C7408.25 25541.9 7423.43 25526.7 7442.1 25526.7H7867.2C7885.86 25526.7 7901.05 25541.9 7901.05 25560.5 7901.05 25579.2 7885.86 25594.4 7867.2 25594.4" fill="#00414c"/><path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M7369.68 26759.3C7544.84 26759.3 7687.36 26616.8 7687.36 26441.6 7687.36 26329.8 7627.48 26224.9 7531.1 26168 7523.32 26163.4 7517.79 26156 7515.54 26147.3 7513.29 26138.5 7514.59 26129.4 7519.19 26121.6 7528.66 26105.6 7549.45 26100.2 7565.55 26109.7 7682.45 26178.8 7755.06 26306 7755.06 26441.6 7755.06 26654.1 7582.18 26827 7369.68 26827 7157.18 26827 6984.3 26654.1 6984.3 26441.6 6984.3 26254.3 7117.99 26094.7 7302.17 26062.2 7304.13 26061.8 7306.12 26061.6 7308.08 26061.6 7316.02 26061.6 7323.75 26064.4 7329.84 26069.6 7337.53 26076 7341.93 26085.5 7341.93 26095.5V26454.2C7341.93 26472.8 7326.74 26488 7308.08 26488 7289.42 26488 7274.23 26472.8 7274.23 26454.2V26138.5C7142.94 26179.6 7052.01 26302.9 7052.01 26441.6 7052.01 26616.8 7194.51 26759.3 7369.68 26759.3" fill="#219a90"/><path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M7267.19 25669.3C7121.88 25681.3 7027.67 25797.9 7019.46 25975.9 7214.19 25963.5 7262.16 25804.5 7267.19 25669.3ZM6994.04 26026.4C6980.35 26026.4 6969.22 26015.2 6969.22 26001.5 6969.22 25768.9 7096.11 25618.6 7292.49 25618.6 7306.18 25618.6 7317.31 25629.7 7317.31 25643.4 7317.31 25890.4 7202.51 26026.4 6994.04 26026.4" fill="#219a90"/><path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M7050.2 25613C6976.66 25556 6879.28 25566.9 6792.87 25642.2 6833.32 25675.4 6876.56 25688.8 6921.53 25681.8 6963.22 25675.3 7006.48 25652.1 7050.2 25613ZM6929.17 25730.8C6860.48 25741.4 6796.4 25717.6 6738.73 25660 6729.05 25650.3 6729.05 25634.5 6738.73 25624.9 6801.71 25561.9 6875.59 25527.2 6946.75 25527.2 7005.68 25527.2 7060.22 25550.6 7104.46 25594.8 7109.15 25599.5 7111.73 25605.7 7111.73 25612.4 7111.73 25619 7109.15 25625.2 7104.46 25629.9 7046.6 25687.8 6987.62 25721.7 6929.17 25730.8" fill="#219a90"/><path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M6732.01 25135.9H6657.73L6694.96 25214.5 6732.01 25135.9ZM6688.67 25287.2 6559.48 25013.8H6599.91L6640.95 25100.5H6748.13L6789.56 25013.8H6829L6701.48 25287.2H6688.67" fill="#00414c"/><path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M7041.53 25074.2C7024.24 25052.4 7002.58 25041.9 6975.32 25041.9 6960.62 25041.9 6947.19 25045.1 6935.42 25051.6 6923.71 25058.1 6914.37 25067.3 6907.65 25078.9 6900.91 25090.6 6897.5 25103.2 6897.5 25116.5 6897.5 25137.1 6904.7 25154.1 6919.51 25168.5 6934.36 25183 6952.56 25190 6975.14 25190 7003.43 25190 7025.12 25179.5 7041.45 25157.8L7044.2 25154.1 7072.14 25171.5 7069.62 25175.7C7063.46 25186 7055.86 25194.6 7047.03 25201.3 7038.25 25208 7027.49 25213.4 7015.04 25217.5 7002.61 25221.6 6989.52 25223.6 6976.13 25223.6 6954.84 25223.6 6935.29 25218.9 6918.03 25209.4 6900.71 25199.9 6886.91 25186.6 6876.99 25169.7 6867.09 25152.8 6862.07 25134.3 6862.07 25114.6 6862.07 25085.3 6872.82 25060 6894.04 25039.5 6915.2 25019.1 6942.16 25008.8 6974.17 25008.8 6994.89 25008.8 7013.68 25012.8 7030.02 25020.9 7046.45 25029 7059.73 25040.5 7069.49 25055.2L7072.27 25059.4 7044.38 25077.8 7041.53 25074.2" fill="#00414c"/><path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M7183.76 25291H7148.68V25218.6H7114.48V25187H7148.68V25013.8H7183.76V25187H7223.52V25218.6H7183.76V25291" fill="#00414c"/><path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M7265.16 25291.3C7260.18 25286.3 7257.65 25280.2 7257.65 25273.1 7257.65 25266.1 7260.18 25260.1 7265.16 25255.1 7270.16 25250 7276.21 25247.5 7283.16 25247.5 7290.22 25247.5 7296.33 25250 7301.32 25255.1 7306.31 25260.1 7308.84 25266.1 7308.84 25273.1 7308.84 25280.2 7306.31 25286.3 7301.32 25291.3 7291.43 25301.3 7275.25 25301.4 7265.16 25291.3" fill="#00414c"/><path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M7265.71 25013.8H7300.78V25218.6H7265.71V25013.8" fill="#00414c"/><path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M7439.85 25083.6 7377.67 25218.6H7339.98L7434.31 25013.8H7445.37L7539.14 25218.6H7501.27L7439.85 25083.6" fill="#00414c"/><path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M7733.03 25168.3C7747.06 25153.9 7753.87 25136.6 7753.87 25115.2 7753.87 25101.4 7750.75 25088.8 7744.59 25077.8 7738.45 25066.8 7729.55 25058 7718.14 25051.6 7695.33 25038.8 7668.1 25038.6 7645.94 25051.6 7634.9 25058.1 7625.98 25067.3 7619.41 25079 7612.83 25090.7 7609.5 25103.2 7609.5 25115.9 7609.5 25128.8 7612.8 25141.3 7619.32 25152.8 7625.82 25164.4 7634.67 25173.5 7645.65 25179.9 7656.64 25186.2 7668.79 25189.5 7681.77 25189.5 7702.26 25189.5 7719.03 25182.5 7733.03 25168.3ZM7752.26 25191.9C7744.21 25200.7 7735.23 25207.7 7725.47 25212.9 7711.86 25220 7696.47 25223.6 7679.72 25223.6 7650.76 25223.6 7625.65 25213.1 7605.06 25192.3 7584.49 25171.4 7574.07 25146 7574.07 25116.7 7574.07 25086.9 7584.4 25061.1 7604.77 25040.3 7625.17 25019.4 7650.14 25008.8 7679 25008.8 7695.23 25008.8 7710.4 25012.2 7724.1 25019 7734.19 25023.9 7743.62 25030.8 7752.26 25039.3V25013.8H7786.98V25218.6H7752.26V25191.9" fill="#00414c"/><path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M7908.65 25291H7873.58V25218.6H7839.37V25187H7873.58V25013.8H7908.65V25187H7948.41V25218.6H7908.65V25291" fill="#00414c"/><path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M7990.05 25291.3C7985.07 25286.3 7982.54 25280.2 7982.54 25273.1 7982.54 25266.1 7985.07 25260.1 7990.05 25255.1 7995.05 25250 8001.1 25247.5 8008.04 25247.5 8015.1 25247.5 8021.22 25250 8026.2 25255.1 8031.2 25260.1 8033.73 25266.1 8033.73 25273.1 8033.73 25280.2 8031.2 25286.3 8026.21 25291.3 8016.32 25301.3 8000.14 25301.4 7990.05 25291.3" fill="#00414c"/><path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M7990.6 25013.8H8025.67V25218.6H7990.6V25013.8" fill="#00414c"/><path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M8230.78 25211.7C8218.72 25219.6 8204.75 25223.6 8189.26 25223.6 8174.16 25223.6 8159.98 25220.1 8147.12 25213.1 8137.91 25208.1 8129.23 25201.2 8121.2 25192.5V25218.6H8086.12V25013.8H8121.2V25090.2C8121.2 25115.4 8122.36 25132.8 8124.62 25142 8128.16 25155.7 8135.59 25167.4 8146.72 25176.7 8157.9 25186 8170.18 25190.6 8184.26 25190.6 8196.3 25190.6 8206.1 25187.6 8213.4 25181.9 8220.68 25176.1 8225.71 25167.3 8228.37 25155.5 8230.12 25148.4 8231.01 25133.7 8231.01 25111.7V25013.8H8266.08V25119.1C8266.08 25146.5 8263.31 25166.3 8257.62 25179.6 8251.87 25193 8242.84 25203.8 8230.78 25211.7" fill="#00414c"/><path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M8488.36 25154.7C8494.35 25144 8497.39 25131.6 8497.39 25117.8 8497.39 25096.7 8490.81 25079.6 8477.84 25066.8 8464.84 25053.9 8447.31 25047.4 8425.74 25047.4 8404.39 25047.4 8386.83 25054 8373.53 25067 8360.09 25080.1 8353.55 25096.2 8353.55 25116.4 8353.55 25129.5 8356.82 25141.9 8363.28 25153.2 8369.71 25164.5 8378.72 25173.4 8390.05 25179.8 8401.42 25186.2 8413.91 25189.5 8427.17 25189.5 8439.84 25189.5 8451.84 25186.3 8462.86 25180.1 8473.82 25173.9 8482.39 25165.3 8488.36 25154.7ZM8496.14 25192.6C8486.67 25202.1 8476.84 25209.3 8466.81 25214.1 8453.58 25220.4 8439.13 25223.6 8423.87 25223.6 8405.49 25223.6 8387.8 25218.8 8371.28 25209.4 8354.73 25199.9 8341.59 25186.8 8332.23 25170.6 8322.87 25154.3 8318.12 25136.5 8318.12 25117.6 8318.12 25098.8 8322.69 25081.3 8331.69 25065.6 8340.7 25049.8 8353.64 25037.1 8370.14 25027.8 8386.6 25018.5 8404.44 25013.8 8423.16 25013.8 8439.12 25013.8 8454.43 25017.2 8468.66 25023.8 8478.9 25028.6 8488.09 25035 8496.12 25042.8 8495.91 25023.8 8493.57 25009.9 8489.18 25001.5 8484.64 24992.8 8476.85 24985.4 8466.04 24979.5 8455.06 24973.6 8441.28 24970.6 8425.11 24970.6 8408.68 24970.6 8394.81 24973.5 8383.88 24979.3 8373.07 24985.1 8364.08 24994.1 8357.18 25006.1L8355.74 25008.6H8317.82L8321.09 25001.5C8328.37 24985.6 8336.51 24973.4 8345.28 24965 8354.09 24956.6 8365.63 24949.8 8379.56 24944.9 8393.35 24940 8409.07 24937.5 8426.28 24937.5 8449.88 24937.5 8470.45 24942.6 8487.4 24952.7 8504.57 24962.9 8516.8 24977.7 8523.73 24996.6 8528.77 25010 8531.21 25030.2 8531.21 25058.4V25218.6H8496.14V25192.6" fill="#00414c"/><path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M8760.69 25190.2C8772.08 25190.2 8784.13 25184.1 8796.53 25172.1L8800.12 25168.6 8823.2 25192.4 8819.6 25195.9C8800.54 25214.3 8781.15 25223.6 8761.95 25223.6 8745.25 25223.6 8731.19 25218.2 8720.17 25207.6 8709.1 25196.8 8703.49 25183.2 8703.49 25167.1 8703.49 25154.7 8707 25143.6 8713.94 25133.9 8720.81 25124.3 8733.13 25115 8751.63 25105.5 8771.2 25095.5 8778.77 25089.3 8781.64 25086 8785.57 25081.2 8787.48 25076 8787.48 25070 8787.48 25062.7 8784.55 25056.5 8778.52 25050.9 8772.36 25045.3 8765.07 25042.6 8756.22 25042.6 8743.41 25042.6 8730.96 25049.3 8719.22 25062.6L8715.47 25066.8 8693.22 25041.6 8695.77 25038.3C8702.9 25029.1 8712.04 25021.8 8722.94 25016.6 8733.78 25011.4 8745.34 25008.8 8757.29 25008.8 8775.44 25008.8 8790.8 25014.9 8802.96 25026.9 8815.13 25038.9 8821.3 25053.8 8821.3 25071.1 8821.3 25083.4 8817.66 25094.7 8810.44 25104.7 8803.2 25114.4 8790.2 25124 8770.67 25134.1 8756.04 25141.6 8746.43 25148.1 8742.07 25153.4 8738.02 25158.3 8736.06 25163.2 8736.06 25168.3 8736.06 25174.1 8738.38 25179 8743.16 25183.5 8748.05 25188 8753.79 25190.2 8760.69 25190.2" fill="#219a90"/><path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M9015.53 25122.7C9015.53 25093.1 9013.63 25081.2 9012.03 25076.4 9008.71 25066.6 9002.56 25058.4 8993.75 25052.4 8984.9 25046.2 8973.87 25043.1 8960.98 25043.1 8948.08 25043.1 8936.87 25046.3 8927.68 25052.6 8918.45 25059 8912.38 25067.3 8909.13 25078.2 8906.87 25086 8905.73 25101 8905.73 25122.7V25218.6H8870.66V25119.8C8870.66 25091.1 8873.88 25070.1 8880.51 25055.5 8887.16 25040.8 8897.39 25029.3 8910.93 25021 8924.4 25012.9 8941.28 25008.8 8961.08 25008.8 8980.91 25008.8 8997.71 25012.9 9011.05 25021.1 9024.45 25029.3 9034.63 25040.7 9041.29 25055.2 9047.84 25069.4 9051.16 25090.6 9051.16 25118V25218.6H9015.53V25122.7" fill="#219a90"/><path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M9162.71 25190.2C9174.11 25190.2 9186.17 25184.1 9198.55 25172.1L9202.16 25168.6 9225.22 25192.4 9221.63 25195.9C9202.58 25214.3 9183.18 25223.6 9163.98 25223.6 9147.28 25223.6 9133.23 25218.2 9122.2 25207.6 9111.14 25196.8 9105.52 25183.2 9105.52 25167.1 9105.52 25154.7 9109.04 25143.5 9115.97 25133.9 9122.84 25124.3 9135.15 25115 9153.65 25105.5 9173.23 25095.5 9180.8 25089.3 9183.67 25086 9187.6 25081.2 9189.52 25076 9189.52 25070 9189.52 25062.7 9186.59 25056.5 9180.55 25050.9 9174.39 25045.3 9167.1 25042.6 9158.25 25042.6 9145.45 25042.6 9132.99 25049.3 9121.25 25062.6L9117.5 25066.8 9095.26 25041.6 9097.8 25038.3C9104.94 25029.1 9114.08 25021.8 9124.97 25016.6 9135.82 25011.4 9147.38 25008.8 9159.32 25008.8 9177.47 25008.8 9192.83 25014.9 9204.99 25026.9 9217.16 25038.9 9223.34 25053.8 9223.34 25071.1 9223.34 25083.4 9219.69 25094.7 9212.47 25104.7 9205.24 25114.4 9192.24 25124 9172.7 25134.1 9158.06 25141.6 9148.45 25148.1 9144.1 25153.4 9140.06 25158.3 9138.08 25163.2 9138.08 25168.3 9138.08 25174.1 9140.41 25179 9145.19 25183.5 9150.08 25188 9155.81 25190.2 9162.71 25190.2" fill="#219a90"/><path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M9332.83 25291H9297.75V25218.6H9263.54V25187H9297.75V25013.8H9332.83V25187H9372.58V25218.6H9332.83V25291" fill="#219a90"/><path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M9568.72 25168.3C9582.56 25154.1 9589.57 25136.3 9589.57 25115.2 9589.57 25101.4 9586.44 25088.8 9580.27 25077.8 9574.14 25066.8 9565.24 25058 9553.83 25051.6 9531.04 25038.8 9503.83 25038.6 9481.62 25051.6 9470.58 25058.1 9461.66 25067.3 9455.09 25079 9448.52 25090.8 9445.18 25103.2 9445.18 25115.9 9445.18 25128.8 9448.49 25141.3 9455.01 25152.8 9461.5 25164.4 9470.36 25173.5 9481.33 25179.9 9492.32 25186.2 9504.48 25189.5 9517.47 25189.5 9537.96 25189.5 9554.72 25182.5 9568.72 25168.3ZM9587.96 25191.9C9579.9 25200.7 9570.92 25207.7 9561.16 25212.9 9547.55 25220 9532.15 25223.6 9515.41 25223.6 9486.46 25223.6 9461.34 25213.1 9440.75 25192.3 9420.18 25171.4 9409.75 25146 9409.75 25116.7 9409.75 25086.9 9420.08 25061.1 9440.47 25040.3 9460.87 25019.4 9485.84 25008.8 9514.68 25008.8 9530.91 25008.8 9546.09 25012.2 9559.8 25019 9569.89 25023.9 9579.31 25030.8 9587.96 25039.3V25013.8H9622.66V25218.6H9587.96V25191.9" fill="#219a90"/><path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M9683.11 25291.3C9678.13 25286.3 9675.6 25280.2 9675.6 25273.1 9675.6 25266.1 9678.13 25260.1 9683.11 25255.1 9688.11 25250 9694.16 25247.5 9701.1 25247.5 9708.16 25247.5 9714.27 25250 9719.27 25255.1 9724.25 25260.1 9726.78 25266.1 9726.78 25273.1 9726.78 25280.2 9724.25 25286.3 9719.27 25291.3 9709.36 25301.3 9693.21 25301.4 9683.11 25291.3" fill="#219a90"/><path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M9683.66 25013.8H9718.73V25218.6H9683.66V25013.8" fill="#219a90"/><path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M9923.84 25211.7C9911.78 25219.6 9897.82 25223.6 9882.33 25223.6 9867.22 25223.6 9853.04 25220.1 9840.17 25213.1 9830.97 25208.1 9822.29 25201.2 9814.25 25192.5V25218.6H9779.19V25013.8H9814.25V25090.2C9814.25 25115.4 9815.41 25132.8 9817.68 25142 9821.22 25155.7 9828.65 25167.4 9839.77 25176.7 9850.95 25186 9863.24 25190.6 9877.32 25190.6 9889.35 25190.6 9899.16 25187.6 9906.47 25181.9 9913.75 25176.1 9918.77 25167.3 9921.42 25155.5 9923.18 25148.4 9924.07 25133.7 9924.07 25111.7V25013.8H9959.15V25119.1C9959.15 25146.5 9956.37 25166.3 9950.68 25179.6 9944.94 25193 9935.91 25203.8 9923.84 25211.7" fill="#219a90"/><path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M10171.4 25168.3C10185.2 25154.1 10192.3 25136.3 10192.3 25115.2 10192.3 25101.4 10189.1 25088.8 10183 25077.8 10176.8 25066.8 10167.9 25058 10156.5 25051.6 10133.7 25038.8 10106.5 25038.6 10084.3 25051.6 10073.3 25058.1 10064.3 25067.3 10057.8 25079 10051.2 25090.8 10047.9 25103.2 10047.9 25115.9 10047.9 25128.8 10051.2 25141.3 10057.7 25152.8 10064.2 25164.4 10073 25173.5 10084 25179.9 10095 25186.2 10107.2 25189.5 10120.2 25189.5 10140.6 25189.5 10157.4 25182.5 10171.4 25168.3ZM10190.6 25191.9C10182.6 25200.7 10173.6 25207.7 10163.8 25212.9 10150.2 25220 10134.8 25223.6 10118.1 25223.6 10089.1 25223.6 10064 25213.1 10043.4 25192.3 10022.9 25171.4 10012.4 25146 10012.4 25116.7 10012.4 25086.9 10022.8 25061.1 10043.2 25040.3 10063.6 25019.4 10088.5 25008.8 10117.4 25008.8 10133.6 25008.8 10148.8 25012.2 10162.5 25019 10172.6 25023.9 10182 25030.8 10190.6 25039.3V25013.8H10225.3V25218.6H10190.6V25191.9" fill="#219a90"/><path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M10469.2 25116.6C10469.2 25103.6 10465.9 25091.2 10459.4 25079.6 10452.9 25068.1 10444.1 25058.9 10433.1 25052.6 10422.1 25046.2 10410 25042.9 10397 25042.9 10376.5 25042.9 10359.8 25049.9 10345.8 25064.2 10331.8 25078.5 10325 25095.9 10325 25117.3 10325 25131.2 10328.1 25143.8 10334.3 25154.8 10340.4 25165.7 10349.3 25174.6 10360.8 25181.1 10372.3 25187.6 10384.6 25190.9 10397.4 25190.9 10409.9 25190.9 10421.8 25187.6 10432.8 25181.1 10443.8 25174.6 10452.7 25165.4 10459.3 25153.7 10465.9 25142 10469.2 25129.5 10469.2 25116.6ZM10399.8 25223.6C10383.6 25223.6 10368.5 25220.3 10354.8 25213.6 10344.8 25208.6 10335.4 25201.8 10326.8 25193.2V25293.8H10291.7V25013.8H10326.8V25040.6C10334.9 25031.7 10343.9 25024.7 10353.6 25019.6 10367.2 25012.4 10382.5 25008.8 10399.1 25008.8 10428 25008.8 10453.1 25019.3 10473.7 25040.2 10494.2 25061 10504.6 25086.4 10504.6 25115.8 10504.6 25145.6 10494.3 25171.3 10474 25192.1 10453.6 25213.1 10428.6 25223.6 10399.8 25223.6" fill="#219a90"/><path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M10554.3 25291.3C10549.3 25286.3 10546.8 25280.2 10546.8 25273.1 10546.8 25266.1 10549.3 25260.1 10554.3 25255.1 10559.3 25250 10565.4 25247.5 10572.3 25247.5 10579.4 25247.5 10585.5 25250 10590.5 25255.1 10595.5 25260.1 10598 25266.1 10598 25273.1 10598 25280.2 10595.5 25286.3 10590.5 25291.3 10580.6 25301.3 10564.4 25301.4 10554.3 25291.3" fill="#219a90"/><path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M10554.9 25013.8H10589.9V25218.6H10554.9V25013.8" fill="#219a90"/><path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M10646.6 25013.8H10681.7V25293.8H10646.6V25013.8" fill="#219a90"/><path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M10737.9 25291.3C10732.9 25286.3 10730.3 25280.2 10730.3 25273.1 10730.3 25266.1 10732.9 25260.1 10737.9 25255.1 10742.9 25250 10748.9 25247.5 10755.8 25247.5 10762.9 25247.5 10769 25250 10774 25255.1 10779 25260.1 10781.5 25266.1 10781.5 25273.1 10781.5 25280.2 10779 25286.3 10774 25291.3 10764.1 25301.3 10747.9 25301.4 10737.9 25291.3" fill="#219a90"/><path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M10738.4 25013.8H10773.5V25218.6H10738.4V25013.8" fill="#219a90"/><path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M10890.9 25291H10855.8V25218.6H10821.6V25187H10855.8V25013.8H10890.9V25187H10930.6V25218.6H10890.9V25291" fill="#219a90"/><path transform="matrix(.1,0,0,-.1,-534,2794.64)" d="M11115.8 25218.6 11054 25076 10990.3 25218.6H10952.9L11035.1 25033.5 10995.4 24942.5H11032.8L11153.4 25218.6H11115.8" fill="#219a90"/>
</svg>
      </div>
      <div class="footer-tagline">
        <strong>Climate Lead Group</strong>
        <p>We help identify sustainable development opportunities and implement solutions for local, national and international entities and countries. Based in Costa Rica, specializing in research, analysis, training, and project implementation across energy, decarbonization, and climate change challenges.</p>
      </div>
    </div>

    <div class="footer-bottom">
      OSTRAM Training &mdash; &copy; Climate Lead Group &mdash; San Jos&eacute;, Costa Rica<br>
      Quality &bull; Integrity &bull; Transparency &bull; Responsibility &bull; Inclusion
    </div>
    <div class="footer-meta">__META__</div>
  </div>
</footer>

</body>
</html>
"""


def render_html(data: Mapping[str, object]) -> str:
    """Render the payload into the self-contained interactive dashboard."""
    payload = json.dumps(data, ensure_ascii=False, sort_keys=True).replace(
        "</", "<\\/"
    )
    title = html_text.escape(
        str(data.get("title") or f"OSTRAM {data['profile_id']} profile report")
    )
    meta = html_text.escape(
        f"Profile {data['profile_id']} — Manifest {data['manifest']} "
        f"— Workspace {data['workspace']}"
    )
    document = _HTML_TEMPLATE
    document = document.replace("__JS__", _DASH_JS)
    document = document.replace("__TITLE__", title)
    document = document.replace("__META__", meta)
    document = document.replace("__PAYLOAD__", payload)
    return document


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
