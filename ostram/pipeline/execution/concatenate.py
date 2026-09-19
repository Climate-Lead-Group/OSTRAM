"""Combine otoole result CSV files into one scenario output table."""

from __future__ import annotations

import argparse
import collections
import csv
import hashlib
import json
import math
import re
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


def _variable_cost_corrections(datafile,d,de,activity,out,source):
 datafile=Path(datafile);scenario=datafile.parent.name.removesuffix('_0')
 source=Path(source)
 assert source.is_file(),source
 explicit={};multiplicity=collections.Counter()
 for r in csv.DictReader(source.open(encoding='utf-8-sig')):
  k=r['REGION'],r['TECHNOLOGY'],str(int(float(r['MODE_OF_OPERATION']))),str(int(float(r['YEAR'])))
  value=float(r['VALUE']);assert k not in explicit or explicit[k]==value
  explicit[k]=value;multiplicity[k]+=1
 def v(p,*k):return d.get(p,{}).get(tuple(map(str,k)),de.get(p,0))
 for k,x in explicit.items():assert abs(v('VariableCost',*k)-x)<=1e-9,('Export/compiled coefficient mismatch',k,x,v('VariableCost',*k))
 omitted=collections.defaultdict(float);expected=collections.defaultdict(float);formerly_double_counted=collections.defaultdict(float);duplicate=collections.defaultdict(float)
 for a in activity.itertuples(index=False):
  r,t,m,y,l=a.REGION,a.TECHNOLOGY,str(a.MODE_OF_OPERATION),str(a.YEAR),a.TIMESLICE;k=r,t,m,y
  energy=a.VALUE*v('YearSplit',l,y);discount=(1+v('DiscountRate',r))**(int(y)-2023+.5)
  expected[r,t,int(y)]+=energy*explicit.get(k,0)*multiplicity[k]
  if k not in explicit:
   assert k not in d['VariableCost'] or v('VariableCost',*k)==0,('Exporter omitted nondefault cost',k)
   omitted[t,y]+=energy*v('VariableCost',*k)/discount
  elif k not in d['VariableCost']:formerly_double_counted[t,y]+=energy*v('VariableCost',*k)/discount
  if multiplicity[k]>1:duplicate[t,y]+=energy*explicit[k]*(multiplicity[k]-1)/discount
 raw={(a.REGION,a.TECHNOLOGY,a.YEAR):a.VALUE for a in pd.read_csv(Path(out)/'AnnualVariableOperatingCost.csv').itertuples()}
 errors=[(k,raw.get(k,0),expected.get(k,0)) for k in raw.keys()|expected.keys() if abs(raw.get(k,0)-expected.get(k,0))>1e-6]
 assert not errors,('Raw VOM export does not match supplied CSV coverage',errors[:10])
 evidence=dict(export_parameter_csv=str(source),unique_explicit_rows=len(explicit),duplicate_identical_keys=sum(n>1 for n in multiplicity.values()),export_annual_cost_check=True,correct_omitted_default_MUSD=math.fsum(omitted.values()),duplicate_export_charge_MUSD=math.fsum(duplicate.values()),compiled_absence_false_omission_MUSD=math.fsum(formerly_double_counted.values()),method='Reconcile actual otoole CSV coverage/multiplicity to compiled coefficients: add only omitted defaults and subtract duplicate identical-row charges. Raw outputs/solution unchanged.')
 evidence.update(export_parameter_csv_sha256=hashlib.sha256(source.read_bytes()).hexdigest(),raw_annual_variable_cost_csv=str(Path(out)/'AnnualVariableOperatingCost.csv'),raw_annual_variable_cost_sha256=hashlib.sha256((Path(out)/'AnnualVariableOperatingCost.csv').read_bytes()).hexdigest())
 for k,x in duplicate.items():omitted[k]-=x
 return omitted,evidence


def reconciled_cost_frames(outputs_folder, datafile, parameter_csv):
    """Reconcile accepted default-VOM/duplicate-export charges, leaving raw files intact.

    Otoole's raw fixed O&M already includes residual stock. Only variable-cost
    coverage is corrected; engineering coefficients and solved activity stay fixed.
    """
    data, defaults = {}, {}
    for match in re.finditer(r'^param(?: default ([^\s]+))?\s*:\s*(\w+)\s*:=\s*\n(.*?);',
                             Path(datafile).read_text(), re.M | re.S):
        default, name, body = match.groups()
        if default is not None:
            defaults[name] = float(default)
        table = {}
        for line in body.splitlines():
            tokens = line.split('#')[0].split()
            if not tokens:
                continue
            key = tuple(tokens[:-1])
            if key in table:
                raise ValueError(f'Duplicate compiled coefficient: {name} {key}')
            table[key] = float(tokens[-1])
        data[name] = table
    activity = pd.read_csv(Path(outputs_folder)/'RateOfActivity.csv')
    correction, evidence = _variable_cost_corrections(
        datafile, data, defaults, activity, outputs_folder, parameter_csv)
    # The accepted full model has one accounting region. Fail closed rather
    # than silently reallocating corrections when this contract changes.
    if set(activity.REGION) != {'GLOBAL'}:
        raise ValueError('Full-model cost reconciliation requires GLOBAL region')
    annual = collections.defaultdict(float)
    for (tech, year), value in correction.items():
        annual[int(year)] += value
    adjusted = {}
    for name in ('TotalDiscountedCost', 'DiscountedCostByTechnology',
                 'DiscountedOperatingCost', 'OperatingCost', 'AnnualVariableOperatingCost'):
        path = Path(outputs_folder)/(name+'.csv')
        if not path.is_file():
            continue
        frame = pd.read_csv(path)
        for idx, row in frame.iterrows():
            year = int(row.YEAR)
            delta = annual[year] if name == 'TotalDiscountedCost' else correction.get((row.TECHNOLOGY,str(year)),0)
            if name in ('OperatingCost', 'AnnualVariableOperatingCost'):
                delta *= (1+data.get('DiscountRate',{}).get(('GLOBAL',), defaults.get('DiscountRate',0)))**(year-2023+.5)
            frame.at[idx,'VALUE'] += delta
        if name != 'TotalDiscountedCost':
            present = {(r.TECHNOLOGY, str(int(r.YEAR))) for r in frame.itertuples()}
            missing = []
            for (tech, year), delta in correction.items():
                if not delta or (tech, year) in present:
                    continue
                if name in ('OperatingCost', 'AnnualVariableOperatingCost'):
                    delta *= (1+data.get('DiscountRate',{}).get(('GLOBAL',), defaults.get('DiscountRate',0)))**(int(year)-2023+.5)
                missing.append(dict(REGION='GLOBAL', TECHNOLOGY=tech, YEAR=int(year), VALUE=delta))
            if missing:
                frame = pd.concat([frame, pd.DataFrame(missing)], ignore_index=True)
        adjusted[name] = frame
    raw = pd.read_csv(Path(outputs_folder)/'TotalDiscountedCost.csv').VALUE.sum()
    evidence.update(raw_cost_MUSD=float(raw), reporting_adjustment_net_MUSD=math.fsum(correction.values()),
                    reconciled_cost_MUSD=float(adjusted['TotalDiscountedCost'].VALUE.sum()),
                    compiled_data_sha256=hashlib.sha256(Path(datafile).read_bytes()).hexdigest())
    return adjusted, evidence

def concatenate_outputs(outputs_folder: Path, output_file: Path, *, datafile=None, parameter_csv=None) -> Path | None:
    """Write ``output_file.csv`` from the non-empty CSVs in ``outputs_folder``."""

    if not outputs_folder.is_dir():
        return None

    if (datafile is None) != (parameter_csv is None):
        raise ValueError('Both compiled data and exact exporter VariableCost CSV are required')
    adjusted, evidence = ({}, None) if datafile is None else reconciled_cost_frames(
        outputs_folder, datafile, parameter_csv)
    frames: list[pd.DataFrame] = []
    parameters: list[str] = []
    allowed = {"Parameter", "VALUE", *SET_COLUMNS}
    for path in sorted(outputs_folder.iterdir(), key=lambda item: item.name):
        if not path.is_file():
            continue
        frame = adjusted[path.stem].copy() if path.stem in adjusted else pd.read_csv(path)
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
    if evidence is not None:
        output_file.with_name(output_file.name+'_cost_reconciliation.json').write_text(
            json.dumps(evidence, indent=2), encoding='utf8')
    return destination


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("outputs_folder", type=Path)
    parser.add_argument("output_file", type=Path)
    parser.add_argument('--datafile', type=Path)
    parser.add_argument('--parameter-csv', type=Path)
    args = parser.parse_args(argv)
    concatenate_outputs(args.outputs_folder.resolve(), args.output_file.resolve(), datafile=args.datafile, parameter_csv=args.parameter_csv)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
