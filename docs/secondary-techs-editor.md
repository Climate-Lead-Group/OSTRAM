# Secondary-technology editor

The secondary-technology editor is an internal preparation facility. It reads
materialized A-O workbooks, writes `Secondary_Techs_Editor.xlsx` into the
selected workspace, and applies reviewed changes back only to mutable scenario
workbooks.

Tracked authority remains in `inputs/`, `config/`, and the scenario workbooks;
an editor workbook is generated state and must not silently become a competing
source.

The editor covers technology activation, parameter values, and optional
interconnection activity limits. Country/region rules come from
`config/preparation/Config_country_codes.yaml`, and bilateral-flow inputs come
from `inputs/preparation/secondary_technologies/`.

Use the canonical workflow:

```powershell
python -m ostram run
```

The implementation modules under
`ostram.pipeline.preparation.secondary_techs` are package internals, not public
script entrypoints. They resolve the explicit project and workspace and never
select files from caller CWD.
