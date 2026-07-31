# Auxiliary tools

OSTRAM retains a small set of utilities used by the maintained data and
execution flows. Historical A1-A6/result-analysis and visualization programs
are not part of the maintained product.

## Configuration loader

**Script:** `t1_confection/Z_AUX_config_loader.py`

This importable module provides cached access to
`Config_country_codes.yaml`. It centralizes country and region codes,
technology mappings, the model start year, renewable fuels, transmission
settings, and related configuration consumed by maintained scripts.

Example:

```python
from Z_AUX_config_loader import get_countries, get_first_year

countries = get_countries()
year = get_first_year()
```

## CSV sorter

**Script:** `t1_confection/Z_AUX_sort_csv.py`

This utility sorts CSV files by their columns to produce deterministic file
ordering. It can be run interactively:

```bash
python t1_confection/Z_AUX_sort_csv.py
```

It can also be imported:

```python
from Z_AUX_sort_csv import sort_csv_files_in_folder

sort_csv_files_in_folder("path/to/csv/folder")
```

## Capital annualization

**Script:** `t1_confection/Z_AUX_capital_annualization_script.py`

B2 invokes this post-processing helper when `annualize_capital: true` in
`Config_MOMF_T1_AB.yaml`. It converts lump-sum capital investment into an
annualized cost stream using a capital recovery factor and writes the
`CapitalInvestmentAnnualized` result into the combined output file.

## Interconnection limits from flows

**Script:** `t1_confection/Z_AUX_D1b_set_trn_limits_from_flows.py`

This helper pre-fills TRN interconnection activity limits in
`Secondary_Techs_Editor.xlsx` from historical bilateral flow data. It belongs
to the D1/D2 manual editing workflow; see {doc}`secondary-techs-editor`.
