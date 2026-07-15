# Legacy batch runners

This directory preserves batch files that targeted former machine-specific checkouts.
They are historical evidence and must not be treated as current commands.

The archived `t1_confection/run_sensitivities.bat` and
`t1_confection/run_directions.bat` hard-code the retired `OSTRAM_clean` path and launch
solver workflows. Their former repository locations now contain fail-closed notices so
an accidental invocation cannot start that obsolete workflow.

The root-level batch files and `RUN_ORDER.md` are separately preserved as historical
final-15 operating records. They also require manual path and configuration review before
use.
