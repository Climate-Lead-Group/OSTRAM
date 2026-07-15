# OSTRAM WS-4 — Session Handover Prompt (paste this whole block into the new session)

You are Claude Code (Opus 4.8) continuing **WS-4: model-quality calibration** for OSTRAM (South Asia OSeMOSYS; solved via `B2_Executing_OG_Model.py` → GLPK `--check` then CPLEX). Work is **non-destructive, in a copy, branch-only, no merge**. Read this file IN FULL, then the docs + memory it points to, then continue. There is **one blocker** to clear (a GLPK `--check` failure) before the final solve — it is fully diagnosed in §4.

---

## 0. Orientation — where this work lives (READ FIRST; avoids a known trap)
- **WS-4 lives ENTIRELY in `C:\Users\luisfernando\Desktop\OSeMOSYS\OSTRAM_ws4_workcopy`** — a self-contained hard copy with **its own `.git`**. Do all WS-4 work there, via the absolute paths in this doc.
- **`OSTRAM_clean` is a DIFFERENT track.** This session's default cwd is `OSTRAM_clean`, a separate git repo currently on branch `phaseB-sensitivity-repro` running the **Phase B sensitivity study**. Its git status, branches, commits, and uncommitted diffs (including changed `*/Outputs/*.csv` solve files) are **UNRELATED to WS-4**. Do not reconcile them, do not commit WS-4 changes there, and do not read that diff as WS-4 state.
- **The WS-4 scripts are real and present** at `t1_confection\A3_process\rules_scripts\` (`apply_base_year_pin.py`, `apply_internal_tx_losses.py`, `apply_interconnector_costs.py`, `apply_internal_transmission.py`, …). If a broad recursive glob like `**/apply_base_year_pin.py` returns "No files found", that is a **glob-scoping artifact, not a missing file** — confirm with `ls` / the absolute path.
- **Memory is not empty:** 12 files at `~/.claude/projects/C--Users-luisfernando-Desktop-OSeMOSYS-OSTRAM-clean/memory/` (incl. `ws4-loss-and-baseyear.md`, `MEMORY.md`). Same glob caveat applies.

## 1. Read first (in `OSTRAM_ws4_workcopy/ws3_transmission_audit/`)
- **`WS4_PREFLIGHT.md`** — the WS-4 plan (loss + base-year pin) and the intended CPLEX sequence.
- **`WS3_calibration_report.md`** — WS-3 context; esp. **§Part 2** (D5 internal transmission) and **§P2.8** (methodology + honest caveats).
- **`WS3_PROMOTION_HANDOFF.md`** — the exact file list + git anchors for the eventual promotion.
- **Memory** (`~/.claude/.../memory/`): `ws4-loss-and-baseyear`, `ws3-internal-tx-decisions`, `ws3-working-copy`, `ostram-env-python-path`, `handoff-cplex-solves`, `cplex-threads-4-laptop`, `cplex-objective-constant-offset`, `verify-reproduction-txt-equality`.

## 2. Environment (critical)
- **WORK IN:** `C:\Users\luisfernando\Desktop\OSeMOSYS\OSTRAM_ws4_workcopy` — the live WS-4 fork. **Folder tiers** (each forked from the last, oldest→newest): `OSTRAM_clean` (git baseline, untouched) → `OSTRAM_ws3_workcopy` (FROZEN @ interconnectors-done) → `OSTRAM_ws3_workcopy_D5` (WS-3 deliverable: interconnectors + D5, all 3 solved, promotion-ready) → `OSTRAM_ws4_workcopy` (WS-4 live). Never edit the older tiers.
- **Env python** (conda NOT on PATH): `C:\Users\luisfernando\anaconda3\envs\OSTRAM-env\python.exe`. Prepend `…\OSTRAM-env;…\Scripts;…\Library\bin` to PATH; set `PYTHONIOENCODING=utf-8`, `PYTHONUTF8=1`.
- **Pipeline** (all under `…\OSTRAM_ws4_workcopy\t1_confection`): `A3_process.py --scenario <s>` (restores `A1_Outputs/_post_a2_snapshot_BAU`, applies rules + the WS-3/WS-4 late stages, delivers to `A1_Outputs/A1_Outputs_<s>`) → `B1_Run_Compiler.py --scenarios "<s>"` (compiles otoole CSVs to `A2_Output_Params/<s>/`) → `B2_Executing_OG_Model.py --scenarios "…"` (GLPK `--check`, then CPLEX; StorageDelayN5 variant, `cplex_threads=4`, `only_main_scenario=False`, config `Config_MOMF_T1_AB.yaml`).
- **B2/CPLEX solves are run by LUIS** in an Anaconda Prompt (`conda activate OSTRAM-env`; ~15 min/scenario). Claude preps A3/B1 + verifies inputs; hand B2 to Luis with the exact command.
- **Scenarios:** `A_Calibrated_BAU`, `B_Optimised_VRE`, `C_Target_VRE`.

## 3. Methodological intent (the WHY)
OSTRAM = 3-scenario South Asia power-system optimisation, 2023–2050. WS-3 (DONE, in `_D5`) calibrated **transmission**: interconnector CapEx sourced from v18 `Interconnector_Params` (a new A3 stage); D5 internal (intra-node) transmission got **per-node residuals** (existing grid = peak×1.2), a **2× RE CapEx premium** (LBNL 2019; exposed YAML slider), and `OperationalLife=40`.

**WS-4** adds two model-quality fixes on top, in `_ws4`:
1. **Internal-transmission losses.** Interconnectors already carry per-corridor losses (OAR 0.93–0.98); internal tx had **0%** (unphysical). Give it a **3% loss** (OAR=0.97; CEA transmission ~3–4%, distribution out of scope). It's a **high-leverage** knob — it taxes *all* throughput, so a 3% loss moved A's system cost **+4.1% (~$91B)**, dwarfing the interconnector/D5 effects. Prime WS-1 sensitivity axis.
2. **Base-year lock.** The 3 scenarios diverge slightly in 2023–2026 because their policy caps bite from year 1; but 2023–2026 is near-term/calibrated history and **should be identical across scenarios**, with divergence only from 2027+. Luis's requirement: **differences ZERO in 2023–2026 — both generation AND capacity** (not just generation — he explicitly rejected an activity-only pin because relaxing the build ceiling would let a scenario build capacity *ahead* of its 2027+ targets). Mechanism: pin every scenario's 2023–2026 to the **calibrated `A_Calibrated_BAU` solve**.

## 4. STATUS + THE BLOCKER

**Loss — DONE + verified.** Stage `apply_internal_tx_losses.py` wired into `A3_process.py` (after `stage_ws3_internal_transmission`); knob `Config_country_codes.yaml → internal_transmission.transmission_loss: 0.03`. Compiled OAR=0.97 for the 6 internal families; interconnectors + DSPTRN + all D5 values intact. **A_Calibrated_BAU solved (loss, no pin) = 2,314,200** (optimal, feasible). B/C not yet solved with final inputs.

**Base-year pin — IMPLEMENTED, inputs verified, but the SOLVE is BLOCKED.** `apply_base_year_pin.py` pins, for 2023–2026, in `A-O_Parametrization.xlsx` "Primary Techs" + "Secondary Techs": `TotalTechnologyAnnualActivityLowerLimit = UpperLimit = A's activity` AND `TotalAnnualMaxCapacityInvestment = TotalAnnualMinCapacityInvestment = A's NewCapacity` (build forced) AND `TotalAnnualMaxCapacity = 9999` (ceiling relaxed). Gated to the 384 real techs via `--tech-csv`. Compiled inputs verified: activity+build **A==B==C**, domain-clean.

**⛔ BLOCKER — GLPK `--check` fails at line 194** (`Residual, Total annual maxcap and mincap investments`) for **`TRNINDNOINDWE` (an interconnector), 2028**:
- OSeMOSYS enforces `TotalAnnualMaxCapacity[y] ≥ ResidualCapacity[y] + Σ_{yy≤y, within OperationalLife} TotalAnnualMinCapacityInvestment[yy]`.
- For interconnectors, `fix_trn_residuals` + `cap_trn_to_residual` (A3 stage-3) maintain a delicate Residual↔MaxCapacity↔MinCapInvest consistency: `TotalAnnualMaxCapacity` is **frozen ≈ residual** (e.g. TRNINDNOINDWE residual = MaxCapacity = **36.72** all years), and commissionings live in `MinCapacityInvestment`.
- The pin **forces `MinCapInvest = A's NewCapacity`** in 2023–2026 (e.g. 1.6 in 2024, 8.4 in 2026). Those forced investments stay "alive" within OperationalLife, so at 2028 `Σ MinCapInvest` (~14.2) makes `Residual+ΣMinInvest ≈ 50.9 > MaxCapacity(36.72)` → check fails. **The base-year build-pin collides with the interconnector residual machinery.**

**Two earlier `--check` bugs already FIXED this session (don't reintroduce):**
- (a) **Out-of-domain (line 179):** the pin flipped `Projection.Mode` on **fuel-keyed rows** (`ELC*01`) in "Secondary Techs" → B1 emitted `TotalAnnualMaxCapacityInvestment[ELCBGDXX01]` → GLPK abort. Fixed by the **`--tech-csv` gate** (only touch the 384 real technologies). Keep it.
- (b) Activity-only pin + `--relax-caps` (an earlier version) let capacity build differ → replaced by the build-pin — which introduced blocker (4).

### THE FIX (do this first; confirm approach with Luis)
The base-year lock must **not build-pin the transmission techs** — they're scenario-independent already (interconnectors frozen by `cap_trn_to_residual`; internal-tx set identically by D5) and their capacity is machinery-managed. Only **generation + mining** techs (`PWR*` non-transmission, `MIN*`) genuinely differ across scenarios in 2023–2026 and need pinning.

**Recommended fix:** restrict `apply_base_year_pin.py` to pin ONLY generation/mining/storage techs — **exclude all transmission** from the pin (interconnectors `TRN[A-Z]{5}[A-Z]{5}`, internal families `RNWTRN/RNWNLI/RNWRPO/PWRTRN/TRNNLI/TRNRPO`, `DSPTRN`). Simplest robust rule: skip any tech whose code contains `TRN` or is in the `cap_trn_to_residual` 18-interconnector allowlist. (Transmission is already identical across scenarios, so excluding it does not break the "zero differences" goal — verify that claim: compare compiled transmission params across A/B/C, expect A==B==C.)
- If build-identity for *generation* still trips the check for a `PWR*` tech with a finite later-year `MaxCapacity` (e.g. `PWRNGSMDVXX` capped at 10), either skip build-pinning that tech or confirm its `MaxCapacity[2027+] ≥ residual + forced build`.
- Open question for Luis: is **activity-pin alone** (generation forced identical, capacity free) acceptable, or is **build-pin** (capacity also identical) required? He asked for build-identity — but the cleanest way that passes the check may be: activity-pin all differing techs + build-pin only where later-year MaxCapacity is open (9999). Resolve this explicitly.

### Re-run sequence after the fix
`A` already has a valid loss solve (2,314,200) = the pin reference. Then:
1. Fix `apply_base_year_pin.py` (+ update its self-test); run `--self-test`.
2. `A3_process.py --scenario <s>` for **A, B, C** (regenerate clean delivered — needed because prior pins edited the delivered files).
3. `apply_base_year_pin.py --input-dir A1_Outputs\A1_Outputs_<s> --from-solve-dir Executables\A_Calibrated_BAU_0\Outputs --tech-csv A2_Output_Params\A_Calibrated_BAU\TECHNOLOGY.csv` for all 3.
4. `B1_Run_Compiler.py --scenarios "A_Calibrated_BAU,B_Optimised_VRE,C_Target_VRE"`.
5. **Hand Luis the B2 command** (below). If GLPK `--check` passes and CPLEX solves, verify the payoff; if a base-year backstop appears, diagnose feasibility.

### B2 command (Luis runs, Anaconda Prompt)
```
conda activate OSTRAM-env
cd /d C:\Users\luisfernando\Desktop\OSeMOSYS\OSTRAM_ws4_workcopy\t1_confection
python B2_Executing_OG_Model.py --scenarios "A_Calibrated_BAU,B_Optimised_VRE,C_Target_VRE"
```

## 5. WS-4 files (persistent, in `_ws4`)
- `t1_confection/A3_process/rules_scripts/apply_internal_tx_losses.py` — loss stage (self-tested). Wired into A3.
- `t1_confection/A3_process/rules_scripts/apply_base_year_pin.py` — base-year pin (self-tested; **needs the §4 fix**). Standalone (post-first-solve), NOT an A3 stage.
- `t1_confection/Config_country_codes.yaml` — `internal_transmission` block (`transmission_loss: 0.03`, `re_capex_multiplier: 2.0`, base costs, `operational_life: 40`) + family `OperationalLife` 20→40.
- `t1_confection/A3_process.py` — WS-3 stages (`stage_ws3_interconnector_costs`, `stage_ws3_internal_transmission`) + WS-4 `stage_ws3_internal_tx_losses`, all after stage 5.
- Everything from `_D5` (interconnector fix, D5 residuals CSV, apply_interconnector_costs.py, apply_internal_transmission.py, v18 template, ws3_transmission_audit/).

## 6. Verification approach (rebuild these checks — the session's verify scripts were in an ephemeral scratchpad)
Read compiled CSVs from `A2_Output_Params/<s>/`. Checks that must pass before handing B2:
- **Loss:** `OutputActivityRatio` = 0.97 for the 6 internal families (RNWTRN/RNWNLI/RNWRPO/PWRTRN/TRNNLI/TRNRPO), all 3 scenarios; interconnectors unchanged (0.93–0.98); DSPTRN=1.0.
- **Pin identity:** for 2023–2026, `TotalTechnologyAnnualActivityUpperLimit == LowerLimit` and **A==B==C**; if build-pinning, `TotalAnnualMaxCapacityInvestment == MinCapacityInvestment` and A==B==C. (For the fixed scope, only the pinned tech set.)
- **Domain-clean:** every TECHNOLOGY in the pinned params ∈ `TECHNOLOGY.csv` (no fuel codes).
- **D5 intact:** internal-tx `CapitalCost` RE 200 / non-RE 100, `OperationalLife`=40, per-node `ResidualCapacity` (RNWTRN/PWRTRN per node; NLI/RPO=0).
- **Interconnectors intact:** 18 corridor CapEx = 380…2800, life 40.
- **Post-solve:** base-year `BCK` 2023 = 0 (feasible); 2023–2026 generation (& capacity) byte-identical across A/B/C; `sum(TotalDiscountedCost)` vs anchors.

## 7. Anchors (sum `TotalDiscountedCost`; StorageDelayN5)
| | A_Calibrated_BAU | B_Optimised_VRE | C_Target_VRE |
|---|--:|--:|--:|
| pre-WS-3 | 2,224,447 | 2,113,985 | 2,158,340 |
| post-interconnector | 2,229,145 | 2,117,860 | 2,163,127 |
| post-D5 | 2,222,829 | 2,118,643 | 2,164,880 |
| **+3% loss (A only, no pin)** | **2,314,200 (+4.11%)** | — | — |

## 8. After WS-4 solves
Compute per-scenario cost deltas (loss + pin) vs anchors; confirm 2023–2026 identical across scenarios; update `WS4_PREFLIGHT.md` + memory. Then **D7 citations are DONE** (in `_D5`); only **promotion** remains (Luis does a clean redo — see `WS3_PROMOTION_HANDOFF.md`). Note the WS-4 changes (loss + base-year pin) are NOT yet in the `_D5` promotion set; fold them in once solved.

## 9. Working rules
Non-destructive (work in `_ws4`); back up before editing workbooks (`*_PRE_*` dirs are auto-made by the stages); self-test every tool; verify each change end-to-end (config → A3 → B1 → compiled) before B2; `git add` specific files only (never `-A`). The 3% loss and the base-year pin are the WS-1 sensitivity knobs — keep them exposed/parameterised.
