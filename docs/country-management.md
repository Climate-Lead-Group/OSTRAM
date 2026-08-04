# Country management

Country and region policy is maintained in
`config/preparation/Config_country_codes.yaml`. The file owns country names,
region mappings, technology defaults, interconnection topology, validation
expectations, and country-template definitions.

## Technology-country matrix

The preparation pipeline builds the technology-country matrix in the selected
workspace. Edit the generated matrix only as runtime input to the next
preparation run; maintained country policy remains in the tracked YAML.

## Validation

The internal country validator checks set membership, required technology
families and parameters, value ranges, demand profiles, storage links, and
referential integrity against `inputs/osemosys_global/`. It is invoked through
the preparation workflow and resolves all sources through `ostram.paths`.

## New-country templates

Define template generation in `Config_country_codes.yaml`, including the new
ISO code, reference country, region suffix, coordinates, and interconnections.
The preparation workflow materializes generated CSV templates beneath the
selected workspace. Review them before deliberately promoting accepted values
into maintained `inputs/` authority.

Run the supported workflow with:

```powershell
python -m ostram run
```

Package implementation modules are internal and must not be launched by file
path. See [configuration](configuration.md) for the YAML schema.
