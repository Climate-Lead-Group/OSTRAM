# Secondary-technology editor

The secondary-technology editor is a manual preparation facility; it is not
part of `python -m ostram run`. It reads the A1 `A-O_*` workbooks under
`<workspace>/preparation/A1_Outputs/`, writes `Secondary_Techs_Editor.xlsx` into
`<workspace>/preparation/`, and applies reviewed changes back to those
workbooks only.

Tracked authority remains in `inputs/`, `config/`, and the scenario workbook;
the editor workbook is generated state and must not become a competing source.

The editor covers technology activation, parameter values, and optional
interconnection activity limits. Country/region rules come from
`config/preparation/Config_country_codes.yaml`, and the source workbooks come
from `inputs/preparation/secondary_technologies/`.

After A1/A2 have populated the workspace, from the repo root:

```powershell
python -B -m ostram.pipeline.preparation.secondary_techs.editor_template
# edit <workspace>/preparation/Secondary_Techs_Editor.xlsx
python -B -m ostram.pipeline.preparation.secondary_techs.transmission_limits   # optional
python -B -m ostram.pipeline.preparation.secondary_techs.update
```

`update` edits the A1 outputs in place after a timestamped backup. A later
`run` that rebuilds A1/A2 overwrites those edits, and the post-A2 snapshot that
A3 restores does not include them unless A2 is re-run afterwards. Run the
modules as shown, never by file path. Module details and gotchas:
`ostram/pipeline/preparation/secondary_techs/README.md`.
