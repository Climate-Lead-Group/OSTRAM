# Country management

Country and region policy is maintained in
`config/preparation/Config_country_codes.yaml`. The file owns country names,
region mappings, technology defaults, interconnection topology, validation
expectations, and country-template definitions (`template_generation`; schema in
[configuration](configuration.md)).

## Technology-country matrix

A1 reads the tracked `inputs/preparation/secondary_technologies/Tech_Country_Matrix.xlsx`.
`python -B -m ostram.pipeline.preparation.tech_country_matrix` regenerates a copy
under `<workspace>/preparation/` with the configured implausible combinations
pre-marked; promoting it into `inputs/` is a write under `inputs/` and requires
approval. Maintained country policy remains in the tracked YAML.

## Adding a country

```powershell
python -m ostram country template --new MMR --ref BGD -i BGD IND   # or from template_generation in the YAML
python -m ostram country merge MMR
python -m ostram country validate MMR
python -m ostram run --skip-pull --compile-only --scenarios A_Calibrated_BAU   # rebuilds A1/A2
```

`country template` clones the reference country's CSV rows into
`<workspace>/preparation/country_templates/<CODE>/`, remapping interconnections
to the listed neighbours, and adds `<CODE>_CF` / `<CODE>_Dem` sheets to the
timeslice workbook after a timestamped backup (under the full profile that is a
write to `inputs/scenarios/OSTRAM_Timeslice_Inputs.xlsx`, which requires
approval). Review the generated CSVs and add the country to `countries:` in the
YAML. `country merge` copies them into the active CSV inputs with `.bak.csv`
backups and renames every post-A2 snapshot so the next run rebuilds A1/A2.
`country validate` checks set membership, required technology families and
parameters, value ranges, demand profiles, storage links, and referential
integrity.

Prefix `--profile <id>` to run the same commands against a profile workspace;
profiles can additionally copy the new country's rows into their prepared
scenario workbook with `scenario sync-country --country <CODE>`
(`examples/unescap/README.md`).

Module details and gotchas are in `ostram/pipeline/preparation/README.md`.
