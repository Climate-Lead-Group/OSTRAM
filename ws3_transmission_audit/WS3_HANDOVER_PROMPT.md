# OSTRAM WS-3 — Session Handover Prompt (paste this whole block into the new session)

You are Claude Code (Opus 4.8) continuing **WS-3: transmission value audit + calibration** for OSTRAM (South Asia OSeMOSYS, CPLEX 22.1.2). Work is **non-destructive, branch-only, no merge**. A prior session finished the interconnector half; you're continuing with internal transmission + wrap-up. Your persistent memory (`MEMORY.md` + the `ws3-*` files) and the deliverables folder already hold the details — read them first, then continue.

## Read first (in the working copy)
- `ws3_transmission_audit/WS3_TASK_LEDGER.md` — full done/pending checklist.
- `ws3_transmission_audit/WS3_value_audit.md` — Phase-1 sourcing audit (the gate).
- `ws3_transmission_audit/WS3_calibration_report.md` — what changed + cost impact.
- Memory: `ws3-working-copy`, `ws3-interconnector-costs-not-live`, `ws3-internal-tx-decisions`, `ostram-env-python-path`, `handoff-cplex-solves`, `cplex-threads-4-laptop`.

## Environment (critical)
- **WORK IN THE COPY:** `C:\Users\luisfernando\Desktop\OSeMOSYS\OSTRAM_ws3_workcopy`. The original `…\OSTRAM_clean` is the **read-only baseline** — never modify it.
- **Env python** (conda NOT on PATH): `C:\Users\luisfernando\anaconda3\envs\OSTRAM-env\python.exe`. For pipeline runs, prepend env dirs to PATH (`…\OSTRAM-env`, `…\Scripts`, `…\Library\bin`) and set `PYTHONIOENCODING=utf-8`, `PYTHONUTF8=1`. A3/B1 shell sub-steps via `sys.executable`, so invoking with this python propagates the env.
- **Pipeline:** `A2_AddTx` → post-A2 snapshot (`A1_Outputs/_post_a2_snapshot_BAU`) → `A3_process.py --scenario <s>` (restores snapshot, applies v18 + rules) → `B1_Run_Compiler.py --scenarios "…"` (writes otoole CSVs to `A2_Output_Params/<s>/`) → **B2** (`B2_Executing_OG_Model.py --scenarios "…"`, CPLEX solve).
- **B2 solves are run by Luis** in an Anaconda Prompt (cplex_threads=4). Claude preps A3/B1 + verifies. Hand off B2 with the exact command, run from the copy root with OSTRAM-env active.
- **Scenarios:** BAU (base), A_Calibrated_BAU, B_Optimised_VRE, C_Target_VRE. C_Target_VRE's A3 needs a **solved A_Calibrated_BAU** first (its `set_vre_targets` reads BAU generation).
- All 3 currently re-run + re-solved with the interconnector changes (anchors A 2,224,447 / B 2,113,985 / C 2,158,340; post +0.18–0.22%, all feasible).

## DONE — interconnectors (complete + verified)
Wiring gap found + fixed: sourced v18 `Interconnector_Params` costs were never consumed (model used legacy distance-computed `OG_csvs` values). Fix = a new **core A3 stage** `stage_ws3_interconnector_costs` (calls `rules_scripts/apply_interconnector_costs.py`) making v18 `Interconnector_Params` the source of truth (CapEx/FOM → Secondary Techs; OperationalLife → Fixed Horizon Parameters). Per-scenario overrides via the `Interconnector_Params.scenario` column (one-row add). Values live in all 3: 18 corridors sourced/research; submarine raised (LK 1250, MV 2800); LK↔MV subsea 1250; 3 corridors added; OperationalLife=40; FOM=1.5%. Verified end-to-end (v18 → A3 → B1 → compiled).

## PENDING — start here

### 1) D5 — INTERNAL (intra-node) transmission
**Locked decisions (Luis):**
- **Uniform cost across all nodes — NO per-node multipliers.** Intra-node transmission is *just accounting*; the study is about interties. Ignore the per-node premiums the research produced (documented, not wanted).
- **Per-node RESIDUALS are wanted** and already computed + desk-checked: `compute_internal_tx_residuals.py` → `outputs/internal_tx_residuals_*.csv`. Per node: `RNWTRN = RE-avail-at-peak × 1.2`; `PWRTRN = peak×1.2 − RE×1.2`; identity holds; NLI/RPO=0. Peak = timeslice-average peak (ACCEPTED as rule-driven — ignore the old "IN_W ~240" spot-check; that was all-India, the 5 India regions sum to ~235 GW). Guards: Bhutan saturates (hydro>peak), Maldives tiny.

**Next actions:**
- **Inject the per-node residuals** into the post-A2 snapshot per full tech code (`RNWTRN<node>` / `PWRTRN<node>` ResidualCapacity), replacing the flat 5 GW; NLI/RPO=0. Mirror the interconnector-wiring approach (idempotent, backed up, self-tested). Verify via A3 → B1 (node-scaled residuals in compiled inputs).
- **Set OperationalLife = 40** for the 6 internal families (RNWTRN/RNWNLI/RNWRPO/PWRTRN/TRNNLI/TRNRPO) — fixes the YAML=20 vs snapshot 50/20 drift. Internal params live in `Config_country_codes.yaml` (A2 reads per family prefix); the `TotalAnnualMaxCapacityInvestment` YAML key is DEAD (not compiled) — leave it.
- **RE-vs-non-RE CapEx — DECIDED (Luis, literature-backed): apply a 2× RE premium.** RNWTRN/RNWNLI/RNWRPO CapitalCost 100→200, FixedCost 4→8 (keep 4% O&M ratio); PWRTRN/TRNNLI/TRNRPO stay 100/4. Expose a named RE-multiplier param defaulted to 2.0 in Config_country_codes.yaml (A2 applies it to the RE families) so WS-1 can slide it (1.5/1.8/2/3×). Basis: per-kW RE transmission premium ~1.5–2.3× (LBNL: gas $44 / wind $70 / solar $103 per kW; solar ~2×, fits this solar/hydro-heavy RE fleet). Do NOT use the 4.5–7× (per-MWh). Apply alongside the residual injection + OperationalLife=40 in ONE A3→B1 re-run, then Luis re-solves (B2). Node-uniform (no per-node multipliers).

### 2) D7 — citations (documentation, no model effect)
Backfill NP↔BD / BT↔BD source text (ADB Bheramara ~$382/kW; WB Nepal–India ETTP; ADB SASEC WP-38). Optionally add a `Source` column to v18 `Interconnector_Params`.

### 3) Promotion
When Luis approves, bring the copy's changes into `OSTRAM_clean` on a **dedicated branch** (branch-only, no merge): `A3_process.py` (+the new stage), `rules_scripts/apply_interconnector_costs.py`, v18 `Interconnector_Params`, the internal-residual injection, `ws3_transmission_audit/`. `git add` specific files only — the copy shows tracked `Executables` as deleted, so **never `git add -A`**.

## Working rules
Non-destructive (work in the copy); pure-calc + emit a desk-check CSV before any injection; timestamped outputs; reversible (backups + JSON change logs); verify each change end-to-end (config/v18 → A3 → B1 → compiled otoole inputs) before handing B2 to Luis. Keep the task tracker + `WS3_TASK_LEDGER.md` current.
