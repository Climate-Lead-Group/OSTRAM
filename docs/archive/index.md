# Historical Record Archive

This archive preserves implementation prompts, run logs, handoffs, ledgers, audit
material, and retired setup guidance. The records were relocated with `git mv` so their
history remains traceable.

Archived records describe the checkout and workflow that existed when they were written.
They may contain former repository names, machine-specific paths, completed solver runs,
or commands that are unsafe for an inspection-only session. Treat them as evidence, not
as current execution instructions. The main documentation and live code are authoritative
for the present workflow.

| Record | Original location | Purpose and status |
|---|---|---|
| [Cleanroom final prompt](cleanroom/CLEANROOM_FINALPROMPT.md) | `/CLEANROOM_FINALPROMPT.md` | Historical rebuild prompt |
| [Cleanroom run log](cleanroom/CLEANROOM_RUNLOG.md) | `/CLEANROOM_RUNLOG.md` | Historical execution and validation log |
| [Cleanroom solve prompt](cleanroom/CLEANROOM_SOLVE_PROMPT.md) | `/CLEANROOM_SOLVE_PROMPT.md` | Historical solver handoff; not an offline command |
| [WACC test prompt](validation/WACC_TEST_PROMPT.md) | `/WACC_TEST_PROMPT.md` | Historical validation prompt |
| [WACC test result](validation/WACC_TEST_RESULT.md) | `/WACC_TEST_RESULT.md` | Preserved result record |
| [WS3 handover prompt](ws3-ws4/WS3_HANDOVER_PROMPT.md) | `/ws3_transmission_audit/WS3_HANDOVER_PROMPT.md` | Historical handoff |
| [WS3 promotion handoff](ws3-ws4/WS3_PROMOTION_HANDOFF.md) | `/ws3_transmission_audit/WS3_PROMOTION_HANDOFF.md` | Historical promotion record |
| [WS3 task ledger](ws3-ws4/WS3_TASK_LEDGER.md) | `/ws3_transmission_audit/WS3_TASK_LEDGER.md` | Historical implementation ledger |
| [WS4 handover prompt](ws3-ws4/WS4_HANDOVER_PROMPT.md) | `/ws3_transmission_audit/WS4_HANDOVER_PROMPT.md` | Historical handoff |
| [WS4 preflight](ws3-ws4/WS4_PREFLIGHT.md) | `/ws3_transmission_audit/WS4_PREFLIGHT.md` | Historical preflight record |
| [Phase-B implementation log](phase-b/PHASE_B_IMPLEMENTATION_LOG.md) | `/t1_confection/sensitivity_expansion/PHASE_B_IMPLEMENTATION_LOG.md` | Historical implementation log |
| [Technical inventory](audits/TECHNICAL_INVENTORY.md) | `/TECHNICAL_INVENTORY.md` | Point-in-time generated repository audit |
| [Legacy Git setup guide](legacy/OSTRAM_Git_Setup_Guide.html) | `/OSTRAM_Git_Setup_Guide.html` | Retired standalone HTML guide |
| [Legacy sensitivity runner](legacy-runners/t1_confection/run_sensitivities.bat) | `/t1_confection/run_sensitivities.bat` | Disabled runner for the retired `OSTRAM_clean` checkout |
| [Legacy directions runner](legacy-runners/t1_confection/run_directions.bat) | `/t1_confection/run_directions.bat` | Disabled runner for the retired `OSTRAM_clean` checkout |
| [Legacy baseline runner](legacy-runners/root/run_baselines.bat) | `/run_baselines.bat` | Disabled machine-specific B2/solver runner |
| [Legacy root sensitivity runner](legacy-runners/root/run_sensitivities.bat) | `/run_sensitivities.bat` | Disabled machine-specific B2/solver runner |
| [Legacy root directions runner](legacy-runners/root/run_directions.bat) | `/run_directions.bat` | Disabled machine-specific B2/solver runner |
