# Auxiliary tools

Auxiliary implementations live inside `ostram.pipeline` and are called by the
canonical pipeline. They are not separate public script entrypoints.

## Configuration and deterministic CSV handling

`ostram.pipeline.preparation.configuration` provides cached access to
`config/preparation/Config_country_codes.yaml`. Preparation stages consume its
country, region, technology, year, renewable-fuel, and transmission settings.

`ostram.pipeline.preparation.sort_csv` supplies deterministic CSV ordering for
maintained preparation operations.

## Capital annualization

`ostram.pipeline.execution.annualization` is reached by B2 when
`annualize_capital: true` in
`config/execution/Config_MOMF_T1_AB.yaml`. It converts the lump-sum
`CapitalInvestment` and `CapitalInvestmentStorage` results into Capital
Recovery Factor (CRF) payment streams and appends them to the combined
inputs/outputs CSV as new rows in two new columns,
`CapitalInvestmentAnnualized` and `CapitalInvestmentStorageAnnualized`. The
new rows carry `Future`, `Scenario`, `REGION`, `TECHNOLOGY` (or `STORAGE`)
and `YEAR`, with every other set column blank, exactly like the
`CapitalInvestment` rows they derive from.

The lifetime and discount rate are the ones the model solved with:
`OperationalLife` / `OperationalLifeStorage` from the input parameters in the
combined file (falling back to the datafile `param default`), and
`DiscountRate` / `DiscountRateStorage` from each scenario's compiled datafile
`Executables/<Scenario>_<Future>/<Scenario>_<Future>.txt`. A non-integer
lifetime is rounded half-up to a whole number of payments. There are no
built-in numeric defaults: a missing datafile, lifetime or rate stops the
step with an error naming what is missing.

The combined file is read in chunks and rewritten as a stream, so the step
never loads the whole campaign table. Running it twice on the same file is
refused. It can also be run on an existing combined file:

```powershell
python -m ostram.pipeline.execution.annualization --input workspace\execution\OSTRAM_StorageDelay_Combined_Inputs_Outputs.csv --executables workspace\execution\Executables
```

`--discount-rate` and `--asset-lifetime` override the model values for every
series when an explicit sensitivity is wanted.

## Secondary-technology helpers

The modules under `ostram.pipeline.preparation.secondary_techs` create and
apply editor workbooks and can pre-fill transmission activity limits. Their
inputs are governed by the preparation configuration and selected workspace;
they do not search caller CWD.

Use `python -m ostram run` for supported workflow execution. Import these
modules only when extending or testing the package; do not launch their source
files by path.
