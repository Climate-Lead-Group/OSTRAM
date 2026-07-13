# WS-4 — Preflight (internal-tx loss + base-year lock)

**Date:** 2026-07-10 · **Copy:** `OSTRAM_ws4_workcopy` (forked from `OSTRAM_ws3_workcopy_D5`) · **No CPLEX run tonight.**
**Status:** LOSS preflighted (inputs verified, no solve needed). BASE-YEAR PIN tool built + self-tested; it awaits the first CPLEX solve to reference.

---

## Change 1 — internal-transmission 3% loss  ✔ PREFLIGHTED

- **What:** `OutputActivityRatio` 1.0 → **0.97** on the 6 internal families (RNWTRN/RNWNLI/RNWRPO, PWRTRN/TRNNLI/TRNRPO), via the AR files' "Demand Techs" Output rows — the same channel that carries the interconnector losses.
- **Stage:** `A3_process/rules_scripts/apply_internal_tx_losses.py` (self-tested), wired into `A3_process.py` after `stage_ws3_internal_transmission`.
- **Knob:** `Config_country_codes.yaml` → `internal_transmission.transmission_loss: 0.03` (CEA transmission ~3–4%, distribution excluded).
- **Verified (A_Calibrated_BAU, A3→B1):** compiled `OutputActivityRatio` internal-tx = **0.97 (3.0% loss)**; interconnectors unchanged (1.7–7.1%); DSPTRN 0%; **all D5 values intact** (RE CapEx 200 / non-RE 100 / life 40 / per-node residuals).

## Change 2 — base-year pin (2023–2026 identical across scenarios)  ⏳ TOOL READY, awaits first solve

- **What:** set `TotalTechnologyAnnualActivityLowerLimit = UpperLimit = reference` for 2023–2026 across Primary+Secondary Techs, so all scenarios reproduce the calibrated run byte-identically in the base window (divergence only 2027+).
- **Tool:** `A3_process/rules_scripts/apply_base_year_pin.py` (self-tested). `--relax-caps` clears the near-term caps (e.g. C_Target_VRE's VRE caps) that would otherwise block the pin.
- **Why not applied tonight:** the reference is the *calibrated A solve*, which shifts once the 3% loss is in → it must be generated from a first CPLEX solve of A-with-loss.

---

## CPLEX sequence (run at the CPLEX machine)

Ready now: `A_Calibrated_BAU` is loss-compiled. Then:

1. **[prep, no CPLEX]** `A3→B1` for `B_Optimised_VRE` (loss).  *(A already done.)*
2. **[CPLEX]** solve `A_Calibrated_BAU` (loss) → this is the pin reference.
3. **[prep]** `A3→B1` for `C_Target_VRE` (loss) — its `set_vre_targets` now reads the A-with-loss solve.
4. **[pin]** for each `A1_Outputs/A1_Outputs_<scenario>`:
   ```
   python A3_process/rules_scripts/apply_base_year_pin.py \
       --input-dir A1_Outputs/A1_Outputs_<scenario> \
       --from-solve Executables/A_Calibrated_BAU_0/Outputs/TotalTechnologyAnnualActivity.csv \
       --relax-caps
   ```
   then **B1 recompile all 3** (the pin edits the delivered `A-O_Parametrization.xlsx`; do **not** re-run A3 afterward or the pin is lost — or wire it as a late A3 stage reading a frozen reference CSV).
5. **[CPLEX]** solve all 3 (pinned + loss).
6. **verify:** 2023–2026 generation byte-identical across scenarios; feasible (base-year backstop 0); report the cost deltas vs the WS-3 anchors.

## Rollback / safety
Every edit backed up next to its target (`*_PRE_INTERNAL_LOSS_*`, `*_PRE_BASEYEAR_PIN_*`). `OSTRAM_ws3_workcopy` (frozen interconnector milestone) and `OSTRAM_ws3_workcopy_D5` (WS-3 deliverable) are untouched.
