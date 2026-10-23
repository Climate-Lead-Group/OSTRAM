"""Serialize maintained aggregate annual received-import policies during compilation.

The policy lives in the scenario's file; flow membership is derived from actual CSV ratios.
Domestic transfers (including all Indian interregional corridors) are excluded.
"""
from pathlib import Path
import math
import pandas as pd
import yaml
from ostram.paths import resolve_paths
from ostram.pipeline.scenarios.rules.apply_internal_tx_losses import read_loss, read_measured_at, BUS_BAR

def render_policy(input_folder, scenario_name):
    """Return the GMPL sets and parameter of the gross import cap of
    ``scenario_name``.

    The policy is read from ``config/scenarios/<scenario>/scenario.yaml``;
    the demand fuels and cross-border TRN flows come from the CSVs in
    ``input_folder``. The demand rows of a country measured at the bus bar hold
    the published requirement x (1 - loss), so its fraction is written as
    fraction / (1 - loss) and the cap stays the fraction of the published
    requirement; a country whose rows are measured at consumption keeps the
    fraction. Returns an empty ``D4ImportCountry`` set when the scenario has no
    policy. Raises ValueError for an invalid policy, a country without demand or
    a country whose demand rows are measured at both points.
    """
    scenario_file = resolve_paths().scenario_config_root / scenario_name / 'scenario.yaml'
    cfg = yaml.safe_load(scenario_file.read_text(encoding='utf8')) if scenario_file.is_file() else {}
    policy = cfg.get('gross_import_cap')
    if not policy:
        return 'set D4ImportCountry := ;\n'
    countries = policy['countries']
    fraction = float(policy['fraction'])
    first, last = int(policy['first_year']), int(policy['last_year'])
    if not math.isfinite(fraction) or not 0 <= fraction <= 1 or first > last:
        raise ValueError('Invalid gross-import policy')
    if len(set(countries)) != len(countries) or not set(countries) <= {'BGD','BTN','IND','LKA','MDV','NPL'}:
        raise ValueError('Invalid/duplicate gross-import country')
    folder = Path(input_folder)
    demand = pd.read_csv(folder/'SpecifiedAnnualDemand.csv')
    ratios = pd.read_csv(folder/'OutputActivityRatio.csv')
    measured = read_measured_at(resolve_paths().scenario_workbook, scenario_name)
    loss = read_loss(resolve_paths().country_config)
    lines = ['set D4ImportCountry := ' + ' '.join(countries) + ';']
    fractions = {}
    for country in countries:
        fuels = sorted(f for f in demand.FUEL.unique() if f.startswith('ELC'+country) and f.endswith('03'))
        if not fuels:
            raise ValueError('Missing demand for '+country)
        points = {measured.get(f) for f in fuels}
        if len(points) != 1 or None in points:
            raise ValueError(f'Demand rows of {country} need one Measured_at; got {sorted(map(str, points))}')
        fractions[country] = fraction / (1.0 - loss) if points == {BUS_BAR} else fraction
        flows = set()
        for r in ratios.itertuples(index=False):
            t = r.TECHNOLOGY
            if t.startswith('TRN') and len(t)==13 and t[3:6]!=t[8:11] and r.FUEL.startswith('ELC'+country) and r.VALUE>0:
                flows.add((t,r.FUEL,int(r.MODE_OF_OPERATION)))
        lines.append('set D4ImportDemand['+country+'] := '+' '.join(fuels)+';')
        lines.append('set D4ImportFlow['+country+'] := '+' '.join('('+','.join(map(str,f))+')' for f in sorted(flows))+';')
    lines.append('param D4ImportFraction :=\n'+'\n'.join(f'{c} {y} {fractions[c]:.17g}' for c in countries for y in range(first,last+1))+'\n;')
    return '\n'.join(lines)+'\n'
