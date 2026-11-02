"""
apply_internal_tx_losses.py
===========================

Late A3 stage stage_internal_tx_losses. Gives the six INTERNAL (intra-node)
transmission families a non-zero
transmission loss, matching how the interconnectors carry losses (via the
output activity ratio in the A-O_AR files), and puts the demand rows measured at
the bus bar on the far side of that loss.

Background
----------
A2 (transmission.py) injects the internal families (RNWTRN/RNWNLI/RNWRPO,
PWRTRN/TRNNLI/TRNRPO) with output activity = 1.0 -> 0% loss, which is not physical. The
interconnectors already carry per-corridor losses (OAR 0.93-0.98) in
A-O_AR_Projections "Demand Techs" (year-indexed Direction=Output rows) and
A-O_AR_Model_Base_Year "Demand Techs" (Value.Fuel.O). This stage sets the
internal families' output activity to (1 - loss), so B1 compiles
OutputActivityRatio = 1 - loss for them.

Loss value: `internal_transmission.transmission_loss` in Config_country_codes.yaml,
an ASSUMPTION explained beside the key: one number for every node (0.03), or a
mapping of node code to number (for example `ARGXX: 0.03`). A mapping names every
node that has internal families and no other; each node's internal families,
bus-bar demand rows and interconnector modes sent from its `ELC<node>03` take
that node's 1 - loss.

Scope: the 6 internal families on the two AR files' "Demand Techs" sheet, and
the interconnector input rows on their "Secondary" sheet. DSPTRN, generators and
storage untouched.

Interconnectors: every interconnector mode whose input fuel is the
`ELC<node>03` of a bus-bar node draws (1 - loss) per unit of activity, because a
bus-bar requirement equals net generation plus imports minus exports and so
already holds the losses that export flows cause. This edits the Direction Input
year cells of A-O_AR_Projections and Value.Fuel.I of A-O_AR_Model_Base_Year.
Modes switched off by a direction rule (0) and modes whose input fuel is measured
at consumption keep their value; output rows and link loss ratios do not change.

Demand: when the configuration has an `internal_transmission` block, every
Demand row whose `Measured_at` is `bus-bar` in the scenario workbook's
Demand_Projection sheet gets SpecifiedAnnualDemand = workbook value x (1 - loss)
in A-O_Demand.xlsx, every year. Rows measured at `consumption` keep the workbook
value. A setup whose scenario workbook has no Demand_Projection sheet (prepared
root workbooks) gives the rows, with their `Measured_at` column, in its root
A-O_Demand.xlsx; the rule reads that tracked file, never the copy it scales. A
cell that already holds the scaled value is left as it is, so a second run
changes nothing; the same holds for the interconnector input cells.

OUTPUT: timestamped backup, in-place edit of the two AR files and A-O_Demand.xlsx,
JSON change log.

USAGE
    python apply_internal_tx_losses.py --input-dir <stage5 or A1_Outputs/...> [--template <workbook>]
    python apply_internal_tx_losses.py --self-test
    python apply_internal_tx_losses.py --input-dir <dir> --restore
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import sys
from datetime import datetime
from pathlib import Path

import yaml

from ostram import yaml_io
from openpyxl import load_workbook, Workbook
from ostram.paths import resolve_paths
from ostram.pipeline.preparation.transmission import TRN_INTERCONNECTION
from ostram.pipeline.scenarios.rules.config_keys import INTERNAL_TRANSMISSION_KEYS, check_keys

PROJ_FILE = "A-O_AR_Projections.xlsx"
BASE_FILE = "A-O_AR_Model_Base_Year.xlsx"
DEMAND_FILE = "A-O_Demand.xlsx"
DEMAND_SHEET = "Demand_Projection"
MEASURED_AT_COL = "Measured_at"
BUS_BAR = "bus-bar"
MEASURED_AT_VALUES = (BUS_BAR, "consumption")
SHEET = "Demand Techs"
LINK_SHEET = "Secondary"
CONFIG_NAME = "Config_country_codes.yaml"

INTERNAL_FAMILIES = ("RNWTRN", "RNWNLI", "RNWRPO", "PWRTRN", "TRNNLI", "TRNRPO")
DEFAULT_LOSS = 0.03
BACKUP_TAG = "_PRE_INTERNAL_LOSS_"

COL_TECH = "Tech"
COL_DIRECTION = "Direction"
COL_VALUE_O = "Value.Fuel.O"
COL_MODE = "Mode.Operation"
COL_FUEL = "Fuel"
COL_FUEL_I = "Fuel.I"
COL_VALUE_I = "Value.Fuel.I"


def _is_internal(t) -> bool:
    return isinstance(t, str) and len(t) == 11 and t[:6] in INTERNAL_FAMILIES


def _is_interconnector(t) -> bool:
    return isinstance(t, str) and bool(TRN_INTERCONNECTION.match(t))


def _values_differ(a, b, tol=1e-9) -> bool:
    if a is None or b is None:
        return a is not b
    try:
        return abs(float(a) - float(b)) > tol
    except (TypeError, ValueError):
        return a != b


# --------------------------------------------------------------------------- config
def resolve_config(explicit):
    """Return the first existing Config_country_codes.yaml among the explicit
    path, $OSTRAM_CONFIG_PATH and ostram/pipeline/Config_country_codes.yaml.

    Raises FileNotFoundError when none exists.
    """
    cands = []
    if explicit:
        cands.append(Path(explicit))
    env = os.environ.get("OSTRAM_CONFIG_PATH")
    if env:
        cands.append(Path(env))
    cands.append(Path(__file__).resolve().parent.parent.parent / CONFIG_NAME)
    for c in cands:
        if c and c.is_file():
            return c
    raise FileNotFoundError("Config_country_codes.yaml not found; tried " + ", ".join(map(str, cands)))


def _check_loss(loss, where="") -> float:
    """Return ``loss`` as a float; raise ValueError when it is outside [0, 1)."""
    loss = float(loss)
    if not (0.0 <= loss < 1.0):
        raise ValueError(f"transmission_loss{where} must be in [0,1); got {loss}")
    return loss


def read_losses(config_path):
    """Return `internal_transmission.transmission_loss` from the config: one float,
    or {node: float} for the per-node mapping form; DEFAULT_LOSS when absent.
    Raise ValueError for a value outside [0, 1)."""
    with open(config_path, "r", encoding="utf-8") as fh:
        data = yaml_io.safe_load(fh) or {}
    blk = data.get("internal_transmission", {}) or {}
    check_keys(blk, INTERNAL_TRANSMISSION_KEYS, config_path, "internal_transmission")
    loss = blk.get("transmission_loss", DEFAULT_LOSS)
    if isinstance(loss, dict):
        return {str(node).strip(): _check_loss(value, f" of {node}") for node, value in loss.items()}
    return _check_loss(loss)


def read_loss(config_path) -> float:
    """Return `internal_transmission.transmission_loss` from the config, or DEFAULT_LOSS when absent; raise ValueError outside [0, 1) or for the per-node mapping form."""
    loss = read_losses(config_path)
    if isinstance(loss, dict):
        raise ValueError(f"{config_path}: transmission_loss is a per-node mapping; this step needs one number")
    return loss


def internal_nodes(*paths) -> set:
    """Return the node codes (chars 6..10) of the internal-family technologies on
    the 'Demand Techs' sheet of the given AR workbooks."""
    nodes = set()
    for path in paths:
        if not Path(path).exists():
            continue
        wb = load_workbook(path, read_only=True, data_only=True)
        try:
            if SHEET not in wb.sheetnames:
                continue
            rows = wb[SHEET].iter_rows(values_only=True)
            hdr = list(next(rows, ()))
            if COL_TECH not in hdr:
                continue
            tc = hdr.index(COL_TECH)
            nodes.update(r[tc][6:11] for r in rows if _is_internal(r[tc]))
        finally:
            wb.close()
    return nodes


def node_output_ratios(losses: dict, nodes) -> dict:
    """Return {node: 1 - loss} for the per-node mapping form; raise ValueError when
    the mapping misses a node of ``nodes`` or names a node outside it."""
    missing, unknown = sorted(set(nodes) - set(losses)), sorted(set(losses) - set(nodes))
    if missing or unknown:
        raise ValueError(
            "transmission_loss mapping must name every node with internal transmission families"
            + (f"; missing: {', '.join(missing)}" if missing else "")
            + (f"; unknown: {', '.join(unknown)}" if unknown else ""))
    return {node: round(1.0 - loss, 6) for node, loss in losses.items()}


def _oar_for(oar, node, where):
    """Return ``oar`` (one number) or its value for ``node`` (per-node mapping)."""
    if not isinstance(oar, dict):
        return oar
    if node not in oar:
        raise ValueError(f"{where}: node {node} has no transmission_loss")
    return oar[node]


def has_loss_block(config_path) -> bool:
    """True when the configuration has an `internal_transmission` block, the
    setups whose demand rows declare where they are measured."""
    with open(config_path, "r", encoding="utf-8") as fh:
        data = yaml_io.safe_load(fh) or {}
    return bool(data.get("internal_transmission"))


# --------------------------------------------------------------------------- demand rows
def resolve_workbook(explicit=None) -> Path:
    """Return the first existing workbook among the explicit path,
    $OSTRAM_TEMPLATE_PATH (the materialized scenario template A3 sets) and the
    setup's scenario workbook.

    Raises FileNotFoundError when none exists.
    """
    cands = []
    if explicit:
        cands.append(Path(explicit))
    env = os.environ.get("OSTRAM_TEMPLATE_PATH")
    if env:
        cands.append(Path(env))
    cands.append(Path(resolve_paths().scenario_workbook))
    for c in cands:
        if c.is_file():
            return c
    raise FileNotFoundError("scenario workbook not found; tried " + ", ".join(map(str, cands)))


def demand_workbook(explicit=None) -> Path:
    """Return the workbook whose Demand_Projection sheet gives the demand rows:
    the scenario workbook (resolve_workbook) or, when it has no such sheet, the
    setup's root A-O_Demand.xlsx (prepared root workbooks), the tracked file and
    never the copy the rule scales."""
    workbook = resolve_workbook(explicit)
    wb = load_workbook(workbook, read_only=True)
    try:
        has_sheet = DEMAND_SHEET in wb.sheetnames
    finally:
        wb.close()
    if has_sheet:
        return workbook
    root = Path(resolve_paths().root_workbooks) / DEMAND_FILE
    return root if root.is_file() else workbook


def read_demand_rows(workbook, scenario=None) -> dict:
    """Return {fuel: {"measured_at": ..., "values": {year: value}}} for the Demand
    rows of the workbook's Demand_Projection sheet.

    A sheet with a `scenario` column gives the BAU rows, replaced by the rows of
    ``scenario``. Raises ValueError when the sheet or the `Measured_at` column is
    absent, or when a Demand row's `Measured_at` is blank or not one of
    MEASURED_AT_VALUES.
    """
    wb = load_workbook(workbook, read_only=True, data_only=True)
    try:
        if DEMAND_SHEET not in wb.sheetnames:
            raise ValueError(f"{workbook}: no {DEMAND_SHEET} sheet")
        rows = wb[DEMAND_SHEET].iter_rows(values_only=True)
        hdr = list(next(rows, ()))
        for col in ("Demand/Share", "Fuel/Tech", MEASURED_AT_COL):
            if col not in hdr:
                raise ValueError(f"{workbook} {DEMAND_SHEET}: no {col} column")
        kind, code, flag = hdr.index("Demand/Share"), hdr.index("Fuel/Tech"), hdr.index(MEASURED_AT_COL)
        scen = hdr.index("scenario") if "scenario" in hdr else None
        years = {}
        for c, h in enumerate(hdr):
            if isinstance(h, int) and 1900 <= h <= 2200:
                years[h] = c
            elif isinstance(h, str) and h.strip().isdigit() and 1900 <= int(h) <= 2200:
                years[int(h)] = c
        base, own = {}, {}
        for r in rows:
            if r[code] is None or r[kind] != "Demand":
                continue
            if scen is not None and r[scen] not in ("BAU", scenario):
                continue
            value = r[flag].strip() if isinstance(r[flag], str) else r[flag]
            if value not in MEASURED_AT_VALUES:
                raise ValueError(
                    f"{workbook} {DEMAND_SHEET} {r[code]}: {MEASURED_AT_COL} is {r[flag]!r}; "
                    f"expected one of {', '.join(MEASURED_AT_VALUES)}")
            target = own if scen is not None and r[scen] != "BAU" else base
            target[str(r[code]).strip()] = {
                "measured_at": value,
                "values": {y: float(r[c]) for y, c in years.items() if r[c] is not None},
            }
        base.update(own)
        return base
    finally:
        wb.close()


def read_measured_at(workbook, scenario=None) -> dict:
    """Return {fuel: Measured_at} for the Demand rows of Demand_Projection (see read_demand_rows)."""
    return {fuel: row["measured_at"] for fuel, row in read_demand_rows(workbook, scenario).items()}


def bus_bar_fuels(workbook, scenario=None) -> set:
    """Return the demand fuels whose Demand_Projection row is measured at the bus bar."""
    return {fuel for fuel, flag in read_measured_at(workbook, scenario).items() if flag == BUS_BAR}


# --------------------------------------------------------------------------- header helper
def _hdr_map(ws):
    return {ws.cell(row=1, column=c).value: c for c in range(1, ws.max_column + 1)
            if ws.cell(row=1, column=c).value is not None}


def _year_cols(ws):
    out = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=1, column=c).value
        if isinstance(v, int) and 1900 <= v <= 2200:
            out[v] = c
        elif isinstance(v, str) and v.strip().split(".")[0].isdigit() and 1900 <= int(float(v)) <= 2200:
            out[int(float(v))] = c
    return out


# --------------------------------------------------------------------------- editors
def edit_projections(path: Path, oar: float) -> dict:
    """A-O_AR_Projections 'Demand Techs': set year cells of internal Output rows to oar."""
    wb = load_workbook(path)
    log = {"file": str(path), "sheet": SHEET, "cells": 0, "rows": 0}
    try:
        if SHEET not in wb.sheetnames:
            log["skipped"] = "sheet absent"; return log
        ws = wb[SHEET]
        h = _hdr_map(ws)
        tc, dc = h.get(COL_TECH), h.get(COL_DIRECTION)
        ycols = _year_cols(ws)
        if tc is None or dc is None or not ycols:
            log["skipped"] = f"missing Tech/Direction/year cols (tc={tc} dc={dc} yrs={len(ycols)})"; return log
        for r in range(2, ws.max_row + 1):
            tech = ws.cell(row=r, column=tc).value
            if not _is_internal(tech):
                continue
            if str(ws.cell(row=r, column=dc).value) != "Output":
                continue
            value = _oar_for(oar, tech[6:11], f"{path} {tech}")
            touched = False
            for y, c in ycols.items():
                cell = ws.cell(row=r, column=c)
                if _values_differ(cell.value, value):
                    cell.value = value
                    log["cells"] += 1
                    touched = True
            if touched:
                log["rows"] += 1
        wb.save(path)
    finally:
        wb.close()
    return log


def edit_base_year(path: Path, oar: float) -> dict:
    """A-O_AR_Model_Base_Year 'Demand Techs': set Value.Fuel.O of internal rows to oar."""
    wb = load_workbook(path)
    log = {"file": str(path), "sheet": SHEET, "cells": 0, "rows": 0}
    try:
        if SHEET not in wb.sheetnames:
            log["skipped"] = "sheet absent"; return log
        ws = wb[SHEET]
        h = _hdr_map(ws)
        tc, vo = h.get(COL_TECH), h.get(COL_VALUE_O)
        if tc is None or vo is None:
            log["skipped"] = f"missing Tech/Value.Fuel.O cols (tc={tc} vo={vo})"; return log
        for r in range(2, ws.max_row + 1):
            tech = ws.cell(row=r, column=tc).value
            if not _is_internal(tech):
                continue
            value = _oar_for(oar, tech[6:11], f"{path} {tech}")
            cell = ws.cell(row=r, column=vo)
            if cell.value is not None and _values_differ(cell.value, value):
                cell.value = value
                log["cells"] += 1
                log["rows"] += 1
        wb.save(path)
    finally:
        wb.close()
    return log


def edit_demand(path: Path, rows: dict, oar: float) -> dict:
    """A-O_Demand 'Demand_Projection': set every year cell of each bus-bar row of
    ``rows`` (read_demand_rows) to the workbook value x oar.

    A cell that already holds that value is left as it is. Raises ValueError when a
    bus-bar fuel has no row, or a cell holds neither the workbook value nor the
    scaled value.
    """
    wb = load_workbook(path)
    log = {"file": str(path), "sheet": DEMAND_SHEET, "cells": 0, "rows": 0, "fuels": []}
    try:
        ws = wb[DEMAND_SHEET]
        h = _hdr_map(ws)
        fc = h.get("Fuel/Tech")
        ycols = _year_cols(ws)
        if fc is None or not ycols:
            raise ValueError(f"{path} {DEMAND_SHEET}: no Fuel/Tech or year columns")
        wanted = {f: row["values"] for f, row in rows.items() if row["measured_at"] == BUS_BAR}
        seen = set()
        for r in range(2, ws.max_row + 1):
            fuel = ws.cell(row=r, column=fc).value
            fuel = fuel.strip() if isinstance(fuel, str) else fuel
            if fuel not in wanted:
                continue
            seen.add(fuel)
            factor = _oar_for(oar, fuel[3:8], f"{path} {fuel}")
            touched = False
            for y, c in ycols.items():
                cell = ws.cell(row=r, column=c)
                if cell.value is None:
                    continue
                if y not in wanted[fuel]:
                    raise ValueError(f"{path} {fuel} {y}: no workbook value to scale")
                stat = wanted[fuel][y]
                scaled = stat * factor
                if not _values_differ(cell.value, scaled, 1e-12 * max(1.0, abs(scaled))):
                    continue
                if _values_differ(cell.value, stat, 1e-12 * max(1.0, abs(stat))):
                    raise ValueError(
                        f"{path} {fuel} {y}: holds {cell.value}, neither the workbook value "
                        f"{stat} nor {stat} x {factor}")
                cell.value = scaled
                log["cells"] += 1
                touched = True
            if touched:
                log["rows"] += 1
        missing = sorted(set(wanted) - seen)
        if missing:
            raise ValueError(f"{path} {DEMAND_SHEET}: no row for bus-bar fuels {', '.join(missing)}")
        log["fuels"] = sorted(wanted)
        wb.save(path)
    finally:
        wb.close()
    return log


def _link_input_value(value, iar, where):
    """Return iar for an interconnector input holding 1; None when it is empty, 0
    (mode switched off) or already iar. Raise ValueError for any other value."""
    if value is None or not _values_differ(value, iar) or not _values_differ(value, 0.0):
        return None
    if not _values_differ(value, 1.0):
        return iar
    raise ValueError(f"{where}: input holds {value}, neither 1, 0 nor {iar}")


def edit_link_projections(path: Path, bus_bar: set, iar: float) -> dict:
    """A-O_AR_Projections 'Secondary': set the year cells of each interconnector
    Input row whose fuel is in ``bus_bar`` to iar (see _link_input_value)."""
    wb = load_workbook(path)
    log = {"file": str(path), "sheet": LINK_SHEET, "cells": 0, "rows": 0, "edited": []}
    try:
        if LINK_SHEET not in wb.sheetnames:
            log["skipped"] = "sheet absent"; return log
        ws = wb[LINK_SHEET]
        h = _hdr_map(ws)
        tc, dc, fc, mc = h.get(COL_TECH), h.get(COL_DIRECTION), h.get(COL_FUEL), h.get(COL_MODE)
        ycols = _year_cols(ws)
        if None in (tc, dc, fc, mc) or not ycols:
            log["skipped"] = f"missing Tech/Direction/Fuel/Mode/year cols (tc={tc} dc={dc} fc={fc} mc={mc})"
            return log
        for r in range(2, ws.max_row + 1):
            tech = ws.cell(row=r, column=tc).value
            fuel = ws.cell(row=r, column=fc).value
            if not _is_interconnector(tech) or str(ws.cell(row=r, column=dc).value) != "Input":
                continue
            if fuel not in bus_bar:
                continue
            value = _oar_for(iar, fuel[3:8], f"{path} {tech} {fuel}")
            cells = 0
            for y, c in ycols.items():
                cell = ws.cell(row=r, column=c)
                new = _link_input_value(cell.value, value, f"{path} {tech} {fuel} {y}")
                if new is not None:
                    cell.value = new
                    cells += 1
            if cells:
                log["cells"] += cells
                log["rows"] += 1
                log["edited"].append({"tech": tech, "mode": ws.cell(row=r, column=mc).value,
                                      "fuel": fuel, "cells": cells})
        wb.save(path)
    finally:
        wb.close()
    return log


def edit_link_base_year(path: Path, bus_bar: set, iar: float) -> dict:
    """A-O_AR_Model_Base_Year 'Secondary': set Value.Fuel.I of each interconnector
    row whose Fuel.I is in ``bus_bar`` to iar (see _link_input_value)."""
    wb = load_workbook(path)
    log = {"file": str(path), "sheet": LINK_SHEET, "cells": 0, "rows": 0, "edited": []}
    try:
        if LINK_SHEET not in wb.sheetnames:
            log["skipped"] = "sheet absent"; return log
        ws = wb[LINK_SHEET]
        h = _hdr_map(ws)
        tc, fc, vc, mc = h.get(COL_TECH), h.get(COL_FUEL_I), h.get(COL_VALUE_I), h.get(COL_MODE)
        if None in (tc, fc, vc, mc):
            log["skipped"] = f"missing Tech/Fuel.I/Value.Fuel.I/Mode cols (tc={tc} fc={fc} vc={vc} mc={mc})"
            return log
        for r in range(2, ws.max_row + 1):
            tech = ws.cell(row=r, column=tc).value
            fuel = ws.cell(row=r, column=fc).value
            if not _is_interconnector(tech) or fuel not in bus_bar:
                continue
            cell = ws.cell(row=r, column=vc)
            value = _oar_for(iar, fuel[3:8], f"{path} {tech} {fuel}")
            new = _link_input_value(cell.value, value, f"{path} {tech} {fuel}")
            if new is not None:
                cell.value = new
                log["cells"] += 1
                log["rows"] += 1
                log["edited"].append({"tech": tech, "mode": ws.cell(row=r, column=mc).value,
                                      "fuel": fuel, "cells": 1})
        wb.save(path)
    finally:
        wb.close()
    return log


# --------------------------------------------------------------------------- backup / restore
def make_backup(d: Path) -> Path:
    """Copy the input folder to a timestamped `_PRE_INTERNAL_LOSS_` sibling and return its path."""
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    b = d.parent / f"{d.name}{BACKUP_TAG}{stamp}"
    if b.exists():
        raise FileExistsError(b)
    shutil.copytree(d, b)
    return b


def restore(d: Path, frm=None) -> Path:
    """Replace the input folder with the given backup, or the newest one, and return the backup used."""
    d = Path(d)
    if frm is None:
        cands = sorted(p for p in d.parent.iterdir()
                       if p.is_dir() and p.name.startswith(f"{d.name}{BACKUP_TAG}"))
        if not cands:
            raise FileNotFoundError(f"no {BACKUP_TAG}* backup by {d}")
        frm = cands[-1]
    if d.is_dir():
        shutil.rmtree(d)
    shutil.copytree(frm, d)
    return frm


# --------------------------------------------------------------------------- orchestration
def run(input_dir, skip_backup=False, config=None, template=None) -> dict:
    """Set the output activity of the internal transmission families to 1 - loss in
    both AR workbooks and, when the configuration has an `internal_transmission`
    block, scale the bus-bar demand rows by 1 - loss and set the input of each
    interconnector mode sent from a bus-bar node to 1 - loss; return the change log.
    With the per-node mapping form each node takes its own 1 - loss."""
    input_dir = Path(input_dir)
    config_path = resolve_config(config)
    loss = read_losses(config_path)
    proj, base = input_dir / PROJ_FILE, input_dir / BASE_FILE
    if not proj.exists():
        raise FileNotFoundError(proj)
    if isinstance(loss, dict):
        oar = node_output_ratios(loss, internal_nodes(proj, base))
    else:
        oar = round(1.0 - loss, 6)
    demand_rows = None
    if has_loss_block(config_path):
        demand_rows = read_demand_rows(demand_workbook(template))
    backup = None if skip_backup else make_backup(input_dir)
    log = {"loss": loss, "oar": oar, "families": list(INTERNAL_FAMILIES),
           "timestamp": datetime.now().isoformat(),
           "backup_dir": str(backup) if backup else None,
           "projections": edit_projections(proj, oar),
           "base_year": edit_base_year(base, oar) if base.exists() else {"skipped": "absent"},
           "demand": (edit_demand(input_dir / DEMAND_FILE, demand_rows, oar) if demand_rows is not None
                      else {"skipped": "no internal_transmission block"})}
    if demand_rows is not None:
        bus_bar = {f for f, row in demand_rows.items() if row["measured_at"] == BUS_BAR}
        log["link_projections"] = edit_link_projections(proj, bus_bar, oar)
        log["link_base_year"] = (edit_link_base_year(base, bus_bar, oar) if base.exists()
                                 else {"skipped": "absent"})
    else:
        log["link_projections"] = log["link_base_year"] = {"skipped": "no internal_transmission block"}
    if backup is not None:
        p = backup.parent / f"{backup.name}_CHANGES.json"
        p.write_text(json.dumps(log, indent=2, default=str))
        log["log_path"] = str(p)
    return log


def _oar_text(oar) -> str:
    """Return the output ratio of a run log as text: the number, or 'per node'."""
    return "per node" if isinstance(oar, dict) else str(oar)


def print_summary(log):
    """Print the loss, output activity ratio and cell counts of a run log."""
    print("=" * 68)
    if isinstance(log["loss"], dict):
        values = ", ".join(f"{v:.3f}" for v in sorted(set(log["loss"].values())))
        print(f"apply_internal_tx_losses  loss per node ({len(log['loss'])} nodes; values {values})")
    else:
        print(f"apply_internal_tx_losses  loss={log['loss']:.3f}  OAR={log['oar']}")
    print("=" * 68)
    for k in ("projections", "base_year"):
        s = log[k]
        if "skipped" in s:
            print(f"  {k}: SKIPPED ({s['skipped']})")
        else:
            print(f"  {k}: {s['cells']} cells / {s['rows']} rows -> {_oar_text(log['oar'])}")
    for k in ("link_projections", "link_base_year"):
        s = log[k]
        if "skipped" in s:
            print(f"  {k}: SKIPPED ({s['skipped']})")
        else:
            print(f"  {k}: {s['cells']} interconnector input cells / {s['rows']} rows -> {_oar_text(log['oar'])}")
    s = log["demand"]
    if "skipped" in s:
        print(f"  demand: SKIPPED ({s['skipped']})")
    else:
        print(f"  demand: {s['cells']} cells / {s['rows']} rows x {_oar_text(log['oar'])} "
              f"(bus-bar fuels: {', '.join(s['fuels'])})")
    print(f"  backup: {log.get('backup_dir', '(skipped)')}")


# --------------------------------------------------------------------------- self-test
def run_self_test() -> int:
    """Run the rule twice on synthetic inputs, print each check and return 0 on success."""
    import tempfile
    print("=" * 68); print("apply_internal_tx_losses.py SELF-TEST"); print("=" * 68)
    ok = True
    with tempfile.TemporaryDirectory() as td:
        td = Path(td)
        (td / CONFIG_NAME).write_text(
            yaml.safe_dump({"internal_transmission": {"transmission_loss": 0.03}}), encoding="utf-8")
        idir = td / "A1_Outputs" / "A1_Outputs_BAU"; idir.mkdir(parents=True)
        # Projections: Demand Techs with internal Output rows (=1) + an interconnector (must NOT change)
        pw = Workbook(); ws = pw.active; ws.title = SHEET
        for c, hh in enumerate(["Mode.Operation", "Tech", "Fuel", "Direction", 2023, 2024, 2025], 1):
            ws.cell(row=1, column=c, value=hh)
        rows = [
            (1, "RNWTRNINDWE", "ELCINDWE02", "Output", 1, 1, 1),   # -> 0.97
            (1, "RNWTRNINDWE", "ELCINDWE00", "Input", 1, 1, 1),    # input stays 1
            (1, "PWRTRNMDVXX", "ELCMDVXX02", "Output", 1, 1, 1),   # -> 0.97
            (1, "TRNBGDXXINDEA", "ELCINDEA04", "Output", 0.983, 0.983, 0.983),  # interconnector, must NOT change
            (1, "PWRSPVINDWE", "ELCINDWE00", "Output", 1, 1, 1),   # non-tx, must NOT change
        ]
        for ri, row in enumerate(rows, 2):
            for c, v in enumerate(row, 1):
                ws.cell(row=ri, column=c, value=v)
        ls = pw.create_sheet(LINK_SHEET)
        ls.append(["Mode.Operation", "Tech", "Fuel", "Direction", 2023, 2024, 2025])
        link_rows = [
            (1, "TRNINDWEMDVXX", "ELCINDWE03", "Input", 1, 1, 1),         # bus-bar sender -> 0.97
            (2, "TRNINDWEMDVXX", "ELCMDVXX03", "Input", 1, 1, 1),         # consumption sender stays 1
            (1, "TRNINDWEMDVXX", "ELCMDVXX04", "Output", 0.95, 0.95, 0.95),  # output unchanged
            (2, "TRNINDWEMDVXX", "ELCINDWE04", "Output", 0.95, 0.95, 0.95),  # output unchanged
        ]
        for row in link_rows:
            ls.append(list(row))
        pw.save(idir / PROJ_FILE); pw.close()
        # Base year: Demand Techs with Value.Fuel.O
        bw = Workbook(); bs = bw.active; bs.title = SHEET
        for c, hh in enumerate(["Tech", "Fuel.O", COL_VALUE_O], 1):
            bs.cell(row=1, column=c, value=hh)
        brows = [("RNWTRNINDWE", "ELCINDWE02", 1), ("PWRTRNMDVXX", "ELCMDVXX02", 1),
                 ("PWRSPVINDWE", "ELCINDWE00", 1)]  # last is non-tx, must NOT change
        for ri, row in enumerate(brows, 2):
            for c, v in enumerate(row, 1):
                bs.cell(row=ri, column=c, value=v)
        lb = bw.create_sheet(LINK_SHEET)
        lb.append(["Mode.Operation", "Fuel.I", COL_VALUE_I, "Tech", "Fuel.O", COL_VALUE_O])
        lb.append([1, "ELCINDWE03", 1, "TRNINDWEMDVXX", "ELCMDVXX04", 1])
        lb.append([2, "ELCMDVXX03", 1, "TRNINDWEMDVXX", "ELCINDWE04", 1])
        bw.save(idir / BASE_FILE); bw.close()
        # Scenario workbook: one bus-bar row, one consumption row
        tw = Workbook(); ts = tw.active; ts.title = DEMAND_SHEET
        ts.append(["Demand/Share", "Fuel/Tech", "Name", MEASURED_AT_COL, 2023, 2024])
        ts.append(["Demand", "ELCINDWE03", "bus bar", BUS_BAR, 100.0, 110.0])
        ts.append(["Demand", "ELCMDVXX03", "consumption", "consumption", 5.0, 6.0])
        tw.save(td / "template.xlsx"); tw.close()
        dw = Workbook(); ds = dw.active; ds.title = DEMAND_SHEET
        ds.append(["Demand/Share", "Fuel/Tech", "Name", "2023", "2024"])
        ds.append(["Demand", "ELCINDWE03", "x", 100.0, 110.0])
        ds.append(["Demand", "ELCMDVXX03", "x", 5.0, 6.0])
        dw.save(idir / DEMAND_FILE); dw.close()

        run(idir, skip_backup=True, config=td / CONFIG_NAME, template=td / "template.xlsx")

        pv = load_workbook(idir / PROJ_FILE, data_only=True).active
        got = {}
        for r in range(2, pv.max_row + 1):
            got[(pv.cell(row=r, column=2).value, pv.cell(row=r, column=4).value)] = pv.cell(row=r, column=5).value
        bv = load_workbook(idir / BASE_FILE, data_only=True).active
        gotb = {bv.cell(row=r, column=1).value: bv.cell(row=r, column=3).value for r in range(2, bv.max_row + 1)}

        def ck(cond, label):
            """Print PASS or FAIL for one check and record a failure."""
            nonlocal ok; print(("  PASS " if cond else "  FAIL ") + label); ok = ok and cond
        ck(got.get(("RNWTRNINDWE", "Output")) == 0.97, "internal RNWTRN Output 1 -> 0.97")
        ck(got.get(("PWRTRNMDVXX", "Output")) == 0.97, "internal PWRTRN Output 1 -> 0.97")
        ck(got.get(("RNWTRNINDWE", "Input")) == 1, "internal Input stays 1")
        ck(got.get(("TRNBGDXXINDEA", "Output")) == 0.983, "interconnector UNCHANGED")
        ck(got.get(("PWRSPVINDWE", "Output")) == 1, "non-tx UNCHANGED")
        ck(gotb.get("RNWTRNINDWE") == 0.97, "base-year internal Value.Fuel.O -> 0.97")
        ck(gotb.get("PWRSPVINDWE") == 1, "base-year non-tx UNCHANGED")
        lv = load_workbook(idir / PROJ_FILE, data_only=True)[LINK_SHEET]
        gotl = {(lv.cell(row=r, column=1).value, lv.cell(row=r, column=4).value):
                [lv.cell(row=r, column=c).value for c in (5, 6, 7)] for r in range(2, lv.max_row + 1)}
        ck(gotl.get((1, "Input")) == [0.97] * 3, "interconnector input from a bus-bar node 1 -> 0.97")
        ck(gotl.get((2, "Input")) == [1] * 3, "interconnector input from a consumption node stays 1")
        ck(gotl.get((1, "Output")) == [0.95] * 3 and gotl.get((2, "Output")) == [0.95] * 3,
           "interconnector outputs UNCHANGED")
        lbv = load_workbook(idir / BASE_FILE, data_only=True)[LINK_SHEET]
        gotlb = {lbv.cell(row=r, column=1).value: (lbv.cell(row=r, column=3).value, lbv.cell(row=r, column=6).value)
                 for r in range(2, lbv.max_row + 1)}
        ck(gotlb.get(1) == (0.97, 1) and gotlb.get(2) == (1, 1),
           "base-year interconnector Value.Fuel.I 0.97 (bus-bar) / 1 (consumption), outputs UNCHANGED")
        dv = load_workbook(idir / DEMAND_FILE, data_only=True).active
        gotd = {dv.cell(row=r, column=2).value: (dv.cell(row=r, column=4).value, dv.cell(row=r, column=5).value)
                for r in range(2, dv.max_row + 1)}
        ck(gotd.get("ELCINDWE03") == (100.0 * 0.97, 110.0 * 0.97), "bus-bar demand x 0.97")
        ck(gotd.get("ELCMDVXX03") == (5.0, 6.0), "consumption demand UNCHANGED")
        log2 = run(idir, skip_backup=True, config=td / CONFIG_NAME, template=td / "template.xlsx")
        ck(log2["projections"]["cells"] == 0 and log2["base_year"]["cells"] == 0
           and log2["demand"]["cells"] == 0 and log2["link_projections"]["cells"] == 0
           and log2["link_base_year"]["cells"] == 0, "idempotent (2nd run 0 cells)")
    print("=" * 68); print("SELF-TEST", "PASSED" if ok else "FAILED"); print("=" * 68)
    return 0 if ok else 1


def main() -> int:
    """Parse the command line, then run the rule, the self-test or a restore; return the exit code."""
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--input-dir", type=Path, default=Path("A1_Outputs/A1_Outputs_BAU"))
    p.add_argument("--config", type=Path, default=None)
    p.add_argument("--template", type=Path, default=None,
                   help="workbook with Demand_Projection (default $OSTRAM_TEMPLATE_PATH, then the setup's)")
    p.add_argument("--skip-backup", action="store_true")
    p.add_argument("--self-test", action="store_true")
    p.add_argument("--restore", action="store_true")
    p.add_argument("--restore-from", type=Path, default=None)
    a = p.parse_args()
    if a.self_test:
        return run_self_test()
    if a.restore or a.restore_from is not None:
        try:
            used = restore(a.input_dir, a.restore_from)
        except Exception as e:
            print(f"ERROR: {e}", file=sys.stderr); return 1
        print(f"Restored {a.input_dir} from {used}"); return 0
    try:
        log = run(a.input_dir, a.skip_backup, a.config, a.template)
    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr); return 1
    print_summary(log)
    return 0


if __name__ == "__main__":
    sys.exit(main())
