# OSTRAM Phase-B — Clipped Baselines, Interconnection Levers & TradeCap Refinement
### Implementation log / handoff record — 2026-07 (branch `validated-baseline-3scenario`, env `OSTRAM-env`)

This document is the paper trail for the work done extending the OSTRAM Phase-B
sensitivity set: fixing `B_Opt_TxCap150`, adding clipped baselines for all three
pathways, building the interconnector-**direction** exercise, and refining the
trade-**volume** cap to a realistic, strictly-labelled 15%. It records *what* was
done, *how* it was implemented, the *file-edit trail*, the *sequencing*, and the
*decisions locked in* so anyone can reproduce or extend it.

---

## 0. Outcome at a glance

**13 scenarios, all solved optimal, zero backstop generation, all validated.**
System cost = `TotalDiscountedCost` sum (M USD). Anchor: `B_Optimised_VRE` = 2,113,984.5.

| Scenario | System cost | vs B_Opt_Clipped |
|---|--:|--:|
| A_Calibrated_BAU | 2,224,446.5 | — |
| A_Calibrated_BAU_Clipped | 2,224,447.8 | (pathway +0.00%) |
| B_Optimised_VRE | 2,113,984.5 | — |
| **B_Opt_Clipped** (reference) | **2,115,081.7** | 0 |
| C_Target_VRE | 2,158,340.4 | — |
| C_Target_VRE_Clipped | 2,149,161.8 | (pathway −0.43%) |
| B_Opt_TradeCap15 (strict) | 2,126,284.1 | +11,202 (+0.53%) |
| B_Opt_SolarCapexHi | 2,122,387.8 | +7,306 (+0.35%) |
| B_Opt_TxCap150 | 2,141,786.5 | +26,705 (+1.26%) |
| B_Opt_IndiaCosts | 2,076,619.9 | −38,462 (−1.82%) |
| B_Opt_IndiaCostsFuel | 2,076,619.9 | −38,462 (−1.82%) |
| B_Opt_DirBidir (validation) | 2,115,081.7 | +0 (reproduces Clipped) |
| B_Opt_DirContractual | 2,120,330.8 | +5,249 (+0.25%) |

### The three interconnection levers (the centrepiece)
Baseline (B_Opt_Clipped): Bangladesh 37.6% self-sufficient / ~62% imported.
(Real-world Bangladesh today ≈ 90% domestic / ~10% imported.)

| Lever | Mechanism | BGD self-sufficiency 2050 | Regional trade 2050 | Cost | Backstop |
|---|---|--:|--:|--:|:--:|
| **Volume** — TradeCap15 | imports ≤ 15% of demand (strict) | 86.7% | 243 TWh | +0.53% | 0 |
| **Capacity** — TxCap150 | corridors ≤ 1.5 × Residual_2023 | 94.5% | 97 TWh | +1.26% | 0 |
| **Direction** — DirContractual | real one-way contracts | 32.1% | 424 TWh | +0.25% | 0 |

Volume and capacity caps pull Bangladesh toward realistic self-sufficiency for <1.3%
system cost; the direction lever alone *deepens* import reliance (it removes BGD's
incentive to over-build domestic generation for export).

### Cost of physical VRE realism (clipped − unclipped, per pathway)
| Pathway | Δ | Note |
|---|--:|---|
| A_Calibrated_BAU | +1 M USD (+0.00%) | BAU barely uses VRE → clip is a no-op (sanity check) |
| B_Optimised_VRE | +1,097 M USD (+0.05%) | trims BGD wind (9.16→3 GW) + LKA solar overbuild |
| C_Target_VRE | −9,179 M USD (−0.43%) | clipping the unphysical 9.16 GW BGD wind + capping the unmeetable MDV NDC floor is *cheaper* |

---

## 1. Workstreams

### 1.1 Fix `B_Opt_TxCap150` (was infeasible pre-solve)

**Problem.** Two distinct data-coherence defects made `glpsol --check` abort:
1. The non-India backstop imports (`TRNNLI*`/`TRNRPO*` for BGD/BTN/LKA/MDV/NPL) were
   zeroed via `TotalAnnualMaxCapacity = 0`, but they carry ~5 GW `ResidualCapacity`
   → violates `Residual ≤ MaxCap` (`check[...] failed`, line 189).
2. `TRNNPLXXBGDXX` had `MinCapInv = 0.04` (2033) while the corridor cap set MaxCap =
   0.06 and Residual = 0.04 → violates `MaxCap ≥ Residual + MinCapInv` (0.06 ≥ 0.08
   is false; line 194). The old `clamp_to_residual` op clamped MinCapInv to Residual,
   not to the available headroom.

**Decisions locked in.**
- Non-India backstops are disabled by **`TotalTechnologyAnnualActivityUpperLimit = 0`
  (activity block)**, not MaxCap=0 — this blocks imports without touching the
  Residual≤MaxCap invariant. (Matches how the trade-cap lever disables backstops.)
- Corridor `MinCapInv` is clamped to **`min(orig, MaxCap − Residual)`** (the headroom),
  guaranteeing `MaxCap ≥ Residual + MinCapInv` with the cap_envelope-equivalent slack.

**Implementation.** `gen_sensitivity_patches.py::gen_txcap150()`:
- backstop loop now emits one AUL=0 edit per backstop (was two MaxCap/MinCapInv edits);
- new per-year `MinCapInv` clamp computed against `headroom = 1.5·Residual − Residual`.

**Result.** Solved dual-simplex optimal; 2,141,786.5 (+1.26%); BGD 94.5% domestic;
zero backstop. (Note: TxCap150's 1.5×Residual rule *blocks* zero-residual corridors —
BTN→BGD, IND-S→LKA — that B_Opt used ~9 and ~8 GW of; a future "TxCap150-Pipeline"
variant could floor those at their evidence pipeline. Not built.)

### 1.2 Clipped baselines for all three pathways

**Goal.** VRE ceilings represent atlas-based physical potential; measure their cost on
each pathway, not just B_Opt.

**Problem (C_Target only).** `C_Target_VRE_Clipped` aborted at `glpsol --check`
(`check[GLOBAL,PWRSPVMDVXX,2039] failed`, line 199 = max producible < activity floor).
Cause: `set_vre_targets` writes an NDC **generation floor** (`ActivityLowerLimit`) sized
to a cap_envelope MaxCap; Maldives solar needs up to 1.73 GW to meet the 33% NDC floor,
but the atlas ceiling is 1 GW → after the clip the floor is physically unproducible.

**Decision locked in.** Where the VRE ceiling clips MaxCap below its pre-clip value,
scale the **activity lower limit only** down by the same ratio `ceil / MaxCap_orig`
(production is linear in capacity, so this re-caps the NDC target at what is physically
buildable and preserves the original cap_envelope headroom). **Never** scale the upper
limit — B_Opt's VRE lower limits are all 0 (→ no-op there) and its upper limits are the
`-1` "unconstrained" sentinel, which scaling would corrupt.

**Implementation.** `apply_patches.py`:
- `apply_ceiling_layer()` now captures pre-clip MaxCap before the `set_flat` overwrite;
- new `scale_activity_to_ceiling()` scales `TotalTechnologyAnnualActivityLowerLimit`.
- Blast radius verified: only `PWRSPVMDVXX` affected (2035-2050), scaled to a flat
  4.222 PJ = the physical output of 1 GW of MDV solar.

**Result.** CalBAU_Clipped ≈ CalBAU (+1 M USD; near-perfect no-op). C_Target_Clipped
= 2,149,161.8 (−0.43% vs C_Target — enforcing physical VRE limits is *cheaper* here).
Per-pathway "cost of VRE realism" block added to `analyse_sensitivity.py`.

### 1.3 Interconnection **direction** exercise (DirBidir + DirContractual)

**Mechanism.** `set_interconnector_direction.py` (authored in the OSTRAM-push clone,
**copied into this repo**). Each corridor `TRN<SRC><DST>` has two flow modes; the script
zeroes the disabled mode's `InputActivityRatio`/`OutputActivityRatio` in
`A-O_AR_Projections.xlsx` + the base-year values in `A-O_AR_Model_Base_Year.xlsx`, and
sets `Projection.Mode = "User defined"` so B1 emits the zeros. Config: per-scenario
`set_interconnector_direction.yaml` (`forward` = SRC→DST, `reverse` = DST→SRC,
`bidirectional`/omit = both modes). Reads the interconnector set from `TECH_TYPES.csv`.

**Research.** Three cited research passes established the real governing flow of each of
the 11 cross-border corridors (see
`reference/interconnector_direction_references.md` for the full table + primary
sources). Verdicts summarised:
- India→Bangladesh (both corridors) — Adani Godda PPA, Bheramara HVDC.
- Bhutan→India (both) — Tala/Chukha/Mangdechhu hydro export.
- Bhutan→Bangladesh — Dorjilung trilateral (planned).
- Nepal→Bangladesh — 2024 tripartite deal.
- India→Sri Lanka — designed bidirectional but import-dominant; reverse uncontracted
  (Adani Mannar/Pooneryn wind cancelled Feb 2025) → `forward`.
- India→Maldives & Sri Lanka→Maldives — purely conceptual; set to the only physically
  plausible direction (island deficit → block impossible island exports).
- **Nepal↔India (both) — genuinely seasonal → left `bidirectional`.**

**Decisions locked in.**
- 9 corridors locked one-way; 2 (Nepal↔India) left bidirectional.
- India-internal corridors are never touched (intra-national transfers ≠ imports).
- Conceptual corridors get the physically-plausible direction even though inactive in
  B_Opt (faithful, non-binding).
- Every direction decision is cited; the two conceptual (Maldives) corridors are flagged
  as *inferred*, not contract-sourced.

**Validation.** `B_Opt_DirBidir` (all bidirectional = script no-op) reproduces
`B_Opt_Clipped` to the decimal (2,115,081.7; identical 376,070 output rows), proving the
machinery is neutral. `B_Opt_DirContractual` compiled with exactly the intended modes
zeroed (verified: 9 corridors × 560 vars = 5,040 fewer LP columns; TRNBGDXXINDEA keeps
only the India→BGD mode).

**Result.** DirContractual = 2,120,330.8 (+0.25%). Blocking Bangladesh's post-2040
export-back to India East makes BGD *more* import-dependent (32.1% domestic) — the
mirror image of the volume/capacity levers.

### 1.4 TradeCap refinement → realistic, strict 15%

**Finding that drove it.** Per-country 2050 import share (demand − domestic gen)/demand:
**Bangladesh +60%** (net importer); Sri Lanka −50%, Nepal −151%, Bhutan −272% (net
exporters); Maldives ~0%; India −3% (self-sufficient). **Bangladesh is the region's only
net importer**, so a region-wide import cap binds on BGD alone.

**Decisions locked in.**
- Replace the arbitrary 30% with a **realistic 15%** (today ≈10%, Bangladesh's own plans
  ≈15-20%, so 30% was above even its ambitions).
- Apply "for all" as a region-wide rule that binds on Bangladesh only (capping the
  exporters is a no-op and risks throttling their exports).
- **Drop the ×1.5 export allowance** so "15%" means a true ≤15% *import* cap. (With the
  allowance the corridor throughput ceiling was 15%-budget + 1.5×export ≈ 20%, and the
  optimiser spent the export headroom on imports → landed at 18.4%. Removing it lands
  imports at 13.3%.) Implemented as a parameter: `export_factor` (default 1.5 preserves
  TradeCap30's solved state; TradeCap15 uses 0.0).

**Result.** Strict TradeCap15 = 2,126,284.1 (+0.53%); **BGD 86.7% domestic / 13.3%
imported**; zero backstop; CO₂ 1,620 Mt (lowest of all scenarios — self-supply displaces
coal-heavy imports). TradeCap30's files remain on disk but are dropped from the report.

---

## 2. File-edit paper trail

### Scripts modified
| File | Change |
|---|---|
| `sensitivity_expansion/gen_sensitivity_patches.py` | `gen_txcap150()`: backstops → AUL=0; corridor `MinCapInv` clamped to `min(orig, MaxCap−Residual)`. `gen_tradecap30()` → **`gen_tradecap(frac, export_factor=1.5)`**; `__main__` calls `gen_tradecap(0.30)` + `gen_tradecap(0.15, export_factor=0.0)`. |
| `sensitivity_expansion/apply_patches.py` | `apply_ceiling_layer()` captures pre-clip MaxCap; new **`scale_activity_to_ceiling()`** scales the NDC activity **lower** limit down by `ceil/MaxCap_orig` (lower-limit only; never the `-1` upper sentinel). |
| `sensitivity_expansion/validate_sensitivity_configs.py` | `coherence_violations()` now also scans **Demand Techs**; `_expected()` covers TxCap150 (AUL on backstops) + TradeCap15; `chk_run3` uses AUL for backstops; `chk_run1` generalised to read `cap_fraction` from patches.json; `SCEN` includes TradeCap15. |
| `analyse_sensitivity.py` | `CLIP_PAIRS` + per-pathway "cost of physical VRE realism" block; `SCEN_ORDER`/`SENSITIVITIES` extended with the clipped + direction scenarios and **TradeCap30→TradeCap15**. |
| `Config_MOMF_T1_AB.yaml` | Toggled `execute_model`/`create_matrix` False→True around each verify pass. **Left at True** (solve mode). |

### Files created
| File | Purpose |
|---|---|
| `A3_process/rules_scripts/set_interconnector_direction.py` | Direction control (copied from OSTRAM-push). |
| `A3_process/rules_scripts/set_interconnector_direction.yaml` | Default no-op (`directions: {}`). |
| `sensitivity_expansion/reference/interconnector_direction_references.md` | Cited per-corridor direction justification. |
| `A3_process/rules_scripts/configs/A_Calibrated_BAU_Clipped/` | patches.json (ceiling-only) + 4 YAMLs. |
| `A3_process/rules_scripts/configs/C_Target_VRE_Clipped/` | patches.json (ceiling-only) + 4 YAMLs. |
| `A3_process/rules_scripts/configs/B_Opt_DirBidir/` | patches.json + 4 YAMLs + set_interconnector_direction.yaml (empty). |
| `A3_process/rules_scripts/configs/B_Opt_DirContractual/` | patches.json + 4 YAMLs + set_interconnector_direction.yaml (9-corridor map). |
| `A3_process/rules_scripts/configs/B_Opt_TradeCap15/` | patches.json (strict, export_factor 0) + 4 YAMLs. |
| `run_directions.bat` | Batch: solve DirBidir + DirContractual → concat → analyse. |

### Files regenerated
- `configs/B_Opt_TxCap150/patches.json` — 40 edits (backstop AUL + MinCapInv headroom).
- `configs/B_Opt_TradeCap15/patches.json` — strict (no export allowance).
- `configs/B_Opt_TradeCap30/patches.json` — regenerated identical to its solved state (export_factor 1.5).
- `A1_Outputs/A1_Outputs_<scenario>/` for each new scenario — patched A-O via apply_patches (timestamped `*_PREPATCH_*` backups retained).
- `sensitivity_comparison.csv`, `sensitivity_report.txt` — 13-scenario refresh.

---

## 3. Sequencing — reproducible per-scenario workflow

Division of labour: **Claude prepped + verified (steps 1-5); Luis ran the CPLEX solves
(step 6) in his Anaconda Prompt**; analysis (step 7) either.

```
# 1. (re)generate patches.json for the computed scenarios
python sensitivity_expansion\gen_sensitivity_patches.py

# 2. build the patched A-O = source A-O + shared VRE ceiling layer
python sensitivity_expansion\apply_patches.py --scenario <SCEN> --source-scenario <BASE>

# 3. DIRECTION scenarios only — overlay the flow-direction edits on the AR files
python A3_process\rules_scripts\set_interconnector_direction.py ^
    --input-dir A1_Outputs\A1_Outputs_<SCEN> ^
    --yaml A3_process\rules_scripts\configs\<SCEN>\set_interconnector_direction.yaml

# 4. compile to the otoole datafile
python B1_Run_Compiler.py --scenarios <SCEN>

# 5. VERIFY before solving — set Config_MOMF_T1_AB.yaml execute_model:False + create_matrix:False
python B2_Executing_OG_Model.py --scenarios <SCEN>      # generates the .txt, no solve
glpsol -m osemosys_fast_preprocessed_storage_delay.txt ^
   -d Executables\<SCEN>_0\Pre_processed_<SCEN>_0_StorageDelayN5_OpenBCK_RMCarefulXLSX.txt ^
   --check --wlp %TEMP%\chk.lp
#   -> expect "Model has been successfully generated", NO "check[...] failed"
#   then RESTORE execute_model:True + create_matrix:True

# 6. SOLVE (Anaconda Prompt)
python B2_Executing_OG_Model.py --scenarios <SCEN>

# 7. analysis (reads existing outputs; no re-solve)
python concat_all_scenarios_2.py
python analyse_sensitivity.py
```

Notes:
- `apply_patches` (Parametrization) and `set_interconnector_direction` (AR files) touch
  **different files** → compose cleanly, order-independent within a scenario build.
- Direction scenarios also carry the VRE ceiling and are measured vs `B_Opt_Clipped`.
- `analyse_sensitivity.py` reads the 1 GB combined CSV — an I/O-bound ~2-5 min read; the
  two concats are the ~15 min tail. (Optional cleanup: repoint it at the ~40 MB
  per-scenario CSVs.)

---

## 4. Decisions locked in (consolidated)

1. **Baseline of comparison** — all sensitivities measured vs `B_Opt_Clipped` (VRE
   ceiling only), isolating each lever from the clip effect.
2. **VRE ceiling** = `min(atlas, B_Opt MaxCap)`, flat all years, applied to every
   scenario; clipped variants exist for all three pathways.
3. **Ceiling ↔ NDC-floor coherence** — clip the activity lower limit (not upper) to the
   physically-buildable level; only affects MDV solar.
4. **TxCap150** — corridor MaxCap = 1.5×Residual_2023; non-India backstops AUL=0;
   MinCapInv clamped to headroom.
5. **Interconnector directions** — 9 corridors locked to real contractual flow, Nepal↔India
   bidirectional (seasonal), India-internal untouched, conceptual corridors set to the
   only plausible (deficit-import) direction; all cited.
6. **TradeCap = strict 15%** — no export allowance (`export_factor=0`), Bangladesh-only
   (region's sole net importer), backstops zeroed. TradeCap30 retained on disk but not
   analysed.
7. **cplex_threads = 4** (4 physical cores; higher oversubscribes and slows).
8. **System cost** = Σ `TotalDiscountedCost`. Raw CPLEX `Objective =` excludes a
   ~157,222 M USD constant term; reported cost = CPLEX objective + 157,222.

---

## 5. Open / not done (candidate follow-ups)
- **TxCap150-Pipeline** — a softer capacity variant that floors zero-residual corridors
  at their evidence pipeline (BTN→BGD 0.75, IND-S→LKA 1.0 GW) instead of blocking them.
- **Node-mapping check** — Dhalkebar–Muzaffarpur physically lands in India's *Eastern*
  region (Bihar), so it arguably belongs to `TRNINDEANPLXX` not `TRNINDNONPLXX`
  (both bidirectional, so no effect on current results).
- **analyse_sensitivity.py speed** — read per-scenario CSVs instead of the 1 GB combined.
- **Written methods/results section** for the report (three levers + BGD framing).
