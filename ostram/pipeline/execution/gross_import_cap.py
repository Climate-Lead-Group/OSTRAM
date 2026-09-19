"""Serialize maintained aggregate annual received-import policies during compilation.

The policy lives in full.yaml; flow membership is derived from actual CSV ratios.
Domestic transfers (including all Indian interregional corridors) are excluded.
"""
from pathlib import Path
import math
import pandas as pd
import yaml
from ostram.paths import resolve_paths

def render_policy(input_folder, scenario_name):
    cfg = yaml.safe_load((resolve_paths().project_root/'config/profiles/full.yaml').read_text(encoding='utf8'))
    policy = cfg.get('scenario_inputs', {}).get(scenario_name, {}).get('gross_import_cap')
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
    lines = ['set D4ImportCountry := ' + ' '.join(countries) + ';']
    for country in countries:
        fuels = sorted(f for f in demand.FUEL.unique() if f.startswith('ELC'+country) and f.endswith('03'))
        if not fuels:
            raise ValueError('Missing demand for '+country)
        flows = set()
        for r in ratios.itertuples(index=False):
            t = r.TECHNOLOGY
            if t.startswith('TRN') and len(t)==13 and t[3:6]!=t[8:11] and r.FUEL.startswith('ELC'+country) and r.VALUE>0:
                flows.add((t,r.FUEL,int(r.MODE_OF_OPERATION)))
        lines.append('set D4ImportDemand['+country+'] := '+' '.join(fuels)+';')
        lines.append('set D4ImportFlow['+country+'] := '+' '.join('('+','.join(map(str,f))+')' for f in sorted(flows))+';')
    lines.append('param D4ImportFraction :=\n'+'\n'.join(f'{c} {y} {fraction:.17g}' for c in countries for y in range(first,last+1))+'\n;')
    return '\n'.join(lines)+'\n'
