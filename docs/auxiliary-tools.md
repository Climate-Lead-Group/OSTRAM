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
`config/execution/Config_MOMF_T1_AB.yaml`. It converts lump-sum capital
investment into an annualized stream and writes the
`CapitalInvestmentAnnualized` result into the combined output.

## Secondary-technology helpers

The modules under `ostram.pipeline.preparation.secondary_techs` create and
apply editor workbooks and can pre-fill transmission activity limits. Their
inputs are governed by the preparation configuration and selected workspace;
they do not search caller CWD.

Use `python -m ostram run` for supported workflow execution. Import these
modules only when extending or testing the package; do not launch their source
files by path.
