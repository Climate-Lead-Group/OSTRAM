# WS-3/WS-4 archived scripts

This directory quarantines historical scripts that could mutate model inputs. They are
preserved for provenance and must not be executed against the current repository.

`set_final_v18_interconnector_values.py` is the original one-shot WS-3 editor. It embeds
an absolute path to a former work copy, creates a backup, and saves the v18 template in
place. Its former path under `ws3_transmission_audit/` now contains a fail-closed notice.

The input-read-only audit and consistency scripts remain under
`ws3_transmission_audit/`; some emit CSV reports, but they do not save model workbooks or
templates in place.
