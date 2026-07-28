# Legacy batch runners

This directory preserves batch files that targeted former machine-specific checkouts.
They are historical evidence and must not be treated as current commands.

The archived `t1_confection/run_sensitivities.bat` and
`t1_confection/run_directions.bat` hard-code the retired `OSTRAM_clean` path and launch
solver workflows. Their former repository locations now contain fail-closed notices so
an accidental invocation cannot start that obsolete workflow.

The archived `root/run_baselines.bat`, `root/run_sensitivities.bat`, and
`root/run_directions.bat` preserve the final-15 machine-specific solver commands. Their
former root locations now contain fail-closed notices.
[`RUN_ORDER.md`](../cleanroom/RUN_ORDER.md) is archived as the historical operating
record; neither it nor the archived batches is a current execution instruction.
