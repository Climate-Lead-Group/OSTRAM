"""Injectable workbook, CSV, configuration, and pickle effects for B1."""

from __future__ import annotations

from collections.abc import Callable, Mapping
from typing import Any
import pickle
import math
from pathlib import Path

import pandas as pd
import yaml


def read_config(
    path: Any,
    *,
    opener: Callable[..., Any] | None = None,
    loader: Callable[[Any], Any] | None = None,
) -> Any:
    """Read the generated YAML config without depending on the console code page."""
    if opener is None:
        opener = open
    if loader is None:
        loader = yaml.safe_load
    with opener(path, "r", encoding="utf-8") as stream:
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


def allocate_country_fuel_costs(tables, config_path, audit_path):
    # Existing compilation YAML embeds this policy; standalone paths remain
    # supported for the accepted focused regression fixtures.
    if config_path is None:
        return tables
    if isinstance(config_path, dict):
        config = config_path
    else:
        config_path = Path(config_path)
        if not config_path.exists():
            return tables
        config = yaml.safe_load(config_path.read_text(encoding="utf-8"))
    if config.get("schema") != "fuel-cost-allocation-v1":
        raise ValueError("Unsupported fuel cost allocation schema")
    if config.get("upstream_base") != "minimum_country_price":
        raise ValueError("Unsupported upstream fuel accounting rule")
    families = set(config["fuel_families"])
    supplier = config["supplier_prefix"]
    consumer = config["consumer_prefix"]
    default = float(config["default_nonfuel_variable_cost"])
    factor = float(config["price_to_model_energy_factor"])
    if not math.isfinite(factor) or factor <= 0:
        raise ValueError("Invalid fuel price energy-basis factor")
    keys = ["REGION", "TECHNOLOGY", "MODE_OF_OPERATION", "YEAR"]
    costs = tables["VariableCost"]
    ratios = tables["InputActivityRatio"]
    selected = ratios[ratios.FUEL.str[:3].isin(families)].copy()
    if not selected.TECHNOLOGY.str.startswith(consumer).all():
        raise ValueError("Unmapped fossil-fuel consumer; review allocation coverage")
    # Demand or a non-MIN producer would need an explicitly supported boundary.
    demand = tables.get("SpecifiedAnnualDemand")
    if demand is not None and ((demand.FUEL.str[:3].isin(families)) & (demand.Value != 0)).any():
        raise ValueError("Direct fossil demand is outside generator fuel allocation")
    output = tables["OutputActivityRatio"]
    supply = output[output.FUEL.str[:3].isin(families)]
    if not supply.TECHNOLOGY.str.startswith(supplier).all() or not (supply.Value == 1).all():
        raise ValueError("Fuel supply boundary is not unit-output MIN supply")

    def unique_values(frame, columns):
        values = {}
        for row in frame.to_dict("records"):
            key = tuple(row[c] for c in columns)
            value = float(row["Value"])
            if not math.isfinite(value):
                raise ValueError(f"Non-finite coefficient: {key}")
            if key in values and values[key] != value:
                raise ValueError(f"Conflicting duplicate coefficient: {key}")
            values[key] = value
        return values

    prices = unique_values(costs, keys)
    physical = unique_values(selected, keys + ["FUEL"])
    bases = {}
    for (region, tech, mode, year, fuel) in physical:
        source = (region, supplier + fuel[:3] + tech[6:9], 1, year)
        if source not in prices:
            raise ValueError(f"Missing country fuel price: {source}")
        price = prices[source] * factor
        if price <= 0:
            raise ValueError(f"Non-positive fuel price: {source}")
        key = (region, fuel[:3], year)
        bases[key] = min(bases.get(key, price), price)
    charge = {}
    audit = []
    for full_key, ratio in physical.items():
        region, tech, mode, year, fuel = full_key
        if ratio < 0:
            raise ValueError(f"Negative fuel input: {full_key}")
        price_tech = supplier + fuel[:3] + tech[6:9]
        source = (region, price_tech, 1, year)
        if source not in prices:
            raise ValueError(f"Missing country fuel price: {source}")
        price = prices[source] * factor
        base = bases[region, fuel[:3], year]
        key = full_key[:4]
        charge[key] = charge.get(key, 0.0) + ratio * (price - base)
        audit.append(dict(REGION=region, TECHNOLOGY=tech, MODE_OF_OPERATION=mode,
                          YEAR=year, FUEL=fuel, price_technology=price_tech,
                          input_ratio=ratio, country_price_MUSD_PJ=prices[source],
                          energy_basis_factor=factor, upstream_base_MUSD_PJ=base,
                          consumer_uplift_MUSD_PJ_activity=ratio*(price-base),
                          fuel_cost_MUSD_PJ_activity=ratio*price,
                          nonfuel_VOM=prices.get(key, default)))
    # One common positive upstream base plus country consumer uplifts equals
    # the intended price exactly when the unchanged fuel balance is tight.
    # No country's low price can be transferred to another country's consumer.
    updated = costs.copy()
    upstream = updated.TECHNOLOGY.str.startswith(supplier) & updated.TECHNOLOGY.str[3:6].isin(families)
    for idx, row in updated.loc[upstream].iterrows():
        key = (row.REGION, row.TECHNOLOGY[3:6], row.YEAR)
        if key not in bases:
            raise ValueError(f"Unmapped upstream fuel family/year: {key}")
        updated.loc[idx, "Value"] = bases[key]
    generated = []
    replaced = set(charge)
    for key, fuel_charge in charge.items():
        row = {c: None for c in costs.columns}
        row.update(dict(zip(keys, key)))
        row.update(PARAMETER="VariableCost", Scenario=costs.Scenario.iloc[0],
                   Value=round(prices.get(key, default) + fuel_charge, 4))
        generated.append(row)
    keep = [tuple(row) not in replaced for row in updated[keys].itertuples(index=False, name=None)]
    result = dict(tables)
    result["VariableCost"] = pd.concat([updated.loc[keep], pd.DataFrame(generated, columns=costs.columns)], ignore_index=True)
    audit_path = Path(audit_path)
    audit_path.parent.mkdir(parents=True, exist_ok=True)
    pd.DataFrame(audit).to_csv(audit_path, index=False)
    return result
