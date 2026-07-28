# WS-3 — Transmission Calibration Report (Phase 3 injection + Phase 4 re-verify)

> **Historical milestone:** This report preserves its 2026-07-09 measurements and
> solve claims. Current authority is accepted correction `d295dcc` and protected
> final-manifest SHA-256
> `778b4706522bc2b29911e74d5b31d24355c84cbe4c0c7d11d1c9680b2ddc9916`.
>
**Status:** Phase 3–4 COMPLETE — all 3 scenarios re-run (A3+B1) and re-solved (B2); verified · original untouched
**Working copy:** `OSTRAM_ws3_workcopy` (original `OSTRAM_clean` untouched) · branch-only, no merge
**Date:** 2026-07-09

---

## 1. What WS-3 changed, and why

Phase 0–1 established (see `WS3_value_audit.md`) that the model's interconnector CapitalCost/FixedCost were **legacy distance-computed values from `OG_csvs_inputs/`** (e.g. BD↔IN_E = 292.487 $/kW), and that the **sourced, cited values in the v18 `Interconnector_Params` sheet were never consumed by the pipeline** — a wiring gap, confirmed by re-running A3 (pre- and post-A3 values identical).

Per your decision ("fix the wiring"), v18 `Interconnector_Params` is now the **single source of truth** for interconnector CapitalCost / FixedCost / OperationalLife, consumed automatically by A3.

## 2. The fix (mechanism)

A new **core A3 stage** — `stage_ws3_interconnector_costs`, calling `rules_scripts/apply_interconnector_costs.py` — runs after the rules-script chain, before delivery. It:

- reads `Interconnector_Params` from the **per-scenario materialized template** (`$OSTRAM_TEMPLATE_PATH`), so it honours the v18 multi-scenario override system automatically;
- writes **CapitalCost, FixedCost** into `Secondary Techs` (year-indexed) and **OperationalLife** into `Fixed Horizon Parameters` (scalar), for every interconnector tech present in the sheet;
- leaves **ResidualCapacity and the Max*/Investment caps** untouched (owned by `fix_trn_residuals` / `cap_trn_to_residual` / `relax_interconnectors`) and **losses / CapacityFactor** untouched (out of WS-3 scope);
- is idempotent, self-tested (`--self-test`), and backs up + logs (when not `--skip-backup`).

**Per-scenario flexibility** (your question): a different interconnector cost for one scenario is a **one-row add** to `Interconnector_Params` — `scenario=<name>, Tech=<TRN…>, Parameter=CapitalCost, value=…` — merged by identity key `(Tech, Parameter)`. No code change; the stage reads the materialized (already-merged) template. Default (BAU rows) applies to all scenarios.

## 3. Values applied (v18 `Interconnector_Params`, now live)

FOM convention: FixedCost = 1.5% × CapitalCost. OperationalLife = 40 yr (all corridors). CapEx $/kW:

| Tech | Corridor | CapEx | Note |
|---|---|--:|---|
| TRNBGDXXINDEA | BD↔IN_E | 380 | sourced (CEA+ADB) |
| TRNBGDXXINDNE | BD↔IN_NE | 250 | sourced |
| TRNBTNXXINDEA | BT↔IN_E | 150 | sourced (CEA+BPC+JICA) |
| TRNBTNXXINDNE | BT↔IN_NE | 180 | sourced |
| TRNINDNONPLXX | NP↔IN_N | 100 | sourced (CEA+WB ETTP) |
| TRNINDEANPLXX | NP↔IN_E | 130 | sourced |
| TRNINDSOLKAXX | LK↔IN_S | **1250** | submarine, raised from 1031 (research 2023 USD) |
| TRNINDEAINDNO | IN_N↔IN_E | 300 | sourced (CEA+POWERGRID) |
| TRNINDNOINDWE | IN_N↔IN_W | 420 | sourced |
| TRNINDEAINDNE | IN_E↔IN_NE | 220 | sourced |
| TRNINDEAINDSO | IN_E↔IN_S | 550 | sourced |
| TRNINDSOINDWE | IN_W↔IN_S | 320 | sourced |
| TRNMDVXXINDSO | MV↔IN_S | **2800** | submarine, raised from 1600 (research @400 MW) |
| TRNNPLXXBGDXX | NP↔BD | 450 | sourced — ADB Bheramara + WB ETTP; **cited in REFS (D7 ✔)** |
| TRNBTNXXBGDXX | BT↔BD | 500 | sourced — ADB SASEC WP-38 + Dorjilung; **cited in REFS (D7 ✔)** |
| TRNINDEAINDWE | IN_E↔IN_W | 691.399 | added to sheet; legacy CEA-consistent value retained |
| TRNINDNEINDNO | IN_NE↔IN_N | 645.703 | added to sheet; legacy CEA-consistent value retained |
| TRNLKAXXMDVXX | LK↔MV | **1250** | added; **repriced subsea** (was 508, mispriced as overhead) |

## 4. Verification

- **A_Calibrated_BAU & B_Optimised_VRE — end-to-end confirmed ✔.** A3 (with the WS-3 stage) + B1 compile complete. Both the delivered workbooks (`Secondary Techs` CapEx/FOM, `Fixed Horizon Parameters` OperationalLife) **and** the compiled otoole inputs (`A2_Output_Params/<scen>/`, the CPLEX-ready parameters B2 solves) carry the exact sourced values — all 18 corridors, CapEx 380…2800, FOM = 1.5%×CapEx, OperationalLife = 40 uniform. Chain **v18 → A3 → B1 → compiled** proven.
- **C_Target_VRE:** A3 gated on a solved BAU (its `set_vre_targets.py` reads `Executables/A_Calibrated_BAU_0` generation). Runs after the BAU B2 solve — see §5.
- **Original repo untouched:** contamination check clean — `OSTRAM_clean` Executables still dated 2026-07-07; all writes landed in the copy via the config's relative paths.
- **Phase 0 re-verify (all 3 solved) ✔:** base year IDENTICAL across scenarios (SpecifiedAnnualDemand, ResidualCapacity, CapitalCost @2023); zero base-year backstop in all three (feasible); scenarios still diverge only via their policy overlays.

## 5. Interim cost impact (subsequently solved)

Baseline objective anchors (pre-change, from Phase 0): A_Calibrated_BAU = 2,224,447 · B_Optimised_VRE = 2,113,985 · C_Target_VRE = 2,158,340.

**Result (A + B solved with corrected costs, 2026-07-09):**

| Scenario | Pre-WS-3 | Post-WS-3 | Delta |
|---|--:|--:|--:|
| A_Calibrated_BAU | 2,224,447 | 2,229,145 | **+4,698 (+0.21%)** |
| B_Optimised_VRE | 2,113,985 | 2,117,860 | **+3,876 (+0.18%)** |
| C_Target_VRE | 2,158,340 | 2,163,127 | **+4,786 (+0.22%)** |

All three remain feasible (zero base-year backstop). Correcting interconnector costs raises total discounted system cost by **~0.2% in every scenario** (A +0.21%, B +0.18%, C +0.22%) — a modest, expected shift: most corridor CapEx rose vs the legacy values (submarine substantially), but interconnectors are a small share of total system cost.

**Completion sequence (interleaved with your B2 solves), all in `OSTRAM_ws3_workcopy`:**
1. B2 solve **A_Calibrated_BAU + B_Optimised_VRE** — compiled inputs are ready in `A2_Output_Params/<scen>/`.
2. I run **A3 + B1 for C_Target_VRE** against the new BAU solve (its VRE targets key off updated BAU generation).
3. B2 solve **C_Target_VRE**.
4. I compute the per-scenario objective delta vs the anchors above and re-run Phase 0.

## 6. Files changed (in the copy only)

- `t1_confection/A3_process.py` — added `stage_ws3_interconnector_costs` + its call after stage 5.
- `t1_confection/A3_process/rules_scripts/apply_interconnector_costs.py` — **new** (the wiring step; self-tested).
- `t1_confection/A3_process/SOASIA_OSeMOSYS_Template_v18.xlsx` — `Interconnector_Params` final values (backup: `*_PRE_WS3_VALUES_*.xlsx`).
- `docs/archive/ws3-ws4/scripts/set_final_v18_interconnector_values.py` — the archived
  original of the logged v18 value editor; its former path under
  `ws3_transmission_audit/` now fails closed.

## 7. Still open at that checkpoint

- **Citations backfill (D7): ✔ DONE (2026-07-09).** NP↔BD (`TRNNPLXXBGDXX`) and BT↔BD (`TRNBTNXXBGDXX`) IEEE citations written into `SOASIA_v18_REFS.xlsx` → `Interconnector_Params` (CapitalCost + FixedCost rows); 0 `[Pending]` cells remain. Backup: `inputs/SOASIA_v18_REFS_PRE_D7_*.xlsx`. Optional remaining polish: a `Source` column in the live v18 `Interconnector_Params`, and REFS rows for the 3 added corridors (already noted in §3). No model effect.
- **D5 (internal transmission):** research complete (per-node premiums + per-kW-vs-per-MWh finding); recommendation pending your return to it. Internal transmission values are unchanged so far (flat $100/$200 not yet applied).
- **LK↔MV subsea value (1250):** low-confidence proxy (no project data); one-cell edit if you want it higher (Maldives-deep ~2800).

---

# Part 2 — D5: Internal (intra-node) transmission calibration

**Status:** COMPLETE — **all three scenarios re-run (A3→B1) and solved (B2)**, optimal + feasible; input, cost, and base-consistency verified (Phase-0 acceptance test 10/12 — the 2 "fails" are the known-benign 2023-generation-differs-by-policy items). **Working copy:** `OSTRAM_ws3_workcopy_D5` (forked from `OSTRAM_ws3_workcopy` at the interconnector milestone; that folder is now the frozen rollback point). **Date:** 2026-07-09. *(Supersedes the "D5 … not yet applied" note in §7.)*

## P2.1 What changed, and why

The six internal-transmission families (RE: RNWTRN / RNWNLI / RNWRPO; non-RE: PWRTRN / TRNNLI / TRNRPO; 10 nodes each = 60 techs) carried pure placeholders: flat `ResidualCapacity` = 5 GW for every node (Maldives ~= India-West), flat `CapitalCost` 100 / `FixedCost` 4, no RE-vs-non-RE split, and `OperationalLife` drifting 50/20. D5 replaces these with:

1. **Per-node `ResidualCapacity`** (existing grid sized at peak × 1.2; `compute_internal_tx_residuals.py`, desk-checked): `RNWTRN<node>` = RE-available-at-peak × 1.2; `PWRTRN<node>` = peak × 1.2 − RE × 1.2; RNWNLI/RNWRPO/TRNNLI/TRNRPO = 0 (repower / new-line build mechanisms). Replaces the flat 5.
2. **RE 2× CapEx premium** (per-kW RE transmission premium ~1.5–2.3×, LBNL 2019): RE families `CapitalCost` 100→**200**, `FixedCost` 4→**8**; non-RE stay 100 / 4. Exposed as a single **live YAML knob** `internal_transmission.re_capex_multiplier` (default 2.0) so WS-1 can slide it (1.5 / 1.8 / 2 / 3×).
3. **`OperationalLife` = 40** for all six families (was 50/20).

Costs are **uniform across nodes** (per Luis: intra-node transmission is accounting; the study is about interties — the per-node cost multipliers the research produced were dropped).

## P2.2 Mechanism — one late A3 stage (why, not snapshot / A2·YAML)

D5 is applied by a new core A3 stage `stage_ws3_internal_transmission` (calls `rules_scripts/apply_internal_transmission.py`), running **after stage 5 and the interconnector-cost stage** — mirroring `apply_interconnector_costs.py` (idempotent, `--self-test`, `--dry-run`, `--restore`, JSON change log). It writes the delivered `A-O_Parametrization.xlsx`: per-node `ResidualCapacity` + RE/non-RE `CapitalCost`/`FixedCost` → **Demand Techs** (year-indexed); `OperationalLife`=40 → **Fixed Horizon Parameters** (scalar).

A late stage — not snapshot injection or A2/YAML wiring — is **required** because internal-tx `OperationalLife` is stamped to 50/20 by the Stage-1 template merge (`3_update_ao_from_extensions.py`); only a post-stage-5 write survives. (Both Stage-3 residual scripts are safe: `cap_trn_to_residual` allowlists the 18 interconnectors; `fix_trn_residuals` reads only *Secondary Techs*, while the internal families live in *Demand Techs*.) The post-A2 snapshot is left **pristine** (flat 5 / 100 / life 20) — it is the rollback point; the stage does the calibration transparently in the delivered workbook.

## P2.3 Per-node `ResidualCapacity` applied (GW, held flat across the horizon)

| Node | RNWTRN | PWRTRN | Note |
|---|--:|--:|---|
| BGDXX | 0.200 | 18.992 | |
| BTNXX | 1.261 | 0.000 | RE saturates (hydro > peak) → PWRTRN floored 0 |
| INDEA | 1.319 | 33.581 | |
| INDNE | 0.596 | 3.334 | |
| INDNO | 15.133 | 66.406 | |
| INDSO | 24.751 | 48.952 | |
| INDWE | 19.552 | 68.075 | |
| LKAXX | 1.104 | 1.557 | |
| MDVXX | 0.001 | 0.295 | small grid |
| NPLXX | 2.446 | 0.330 | |

RNWNLI / RNWRPO / TRNNLI / TRNRPO = 0 at every node.

## P2.4 Verification (inputs)

- Chain **config/residuals → A3 (late stage) → B1 → compiled otoole inputs** proven for A_Calibrated_BAU + B_Optimised_VRE: compiled `ResidualCapacity` per-node (RNWTRN/PWRTRN; NLI/RPO=0), `CapitalCost` RE 200 / non-RE 100, `FixedCost` RE 8 / non-RE 4, `OperationalLife` 40 (all six).
- **A == B** for all four internal-tx params in the compiled inputs (0 value mismatches — internal tx is scenario-independent, as expected).
- **Interconnectors intact** — all 18 corridor CapEx still match the sourced values (380 … 2800) and life=40; untouched by the D5 stage (regression-checked in the compiled inputs).
- **Snapshot pristine** — `_post_a2_snapshot_BAU` still flat 5 / 100 / life 20 (reversible; rollback point).

## P2.5 Cost impact

All solved runs OPTIMAL (CPLEX dual simplex) and feasible (base-year `BCK` = 0). Metric: sum `TotalDiscountedCost` (same as the WS-3 anchors). Interconnector anchors = A 2,229,145 · B 2,117,860 · C 2,163,127; pre-WS-3 anchors = A 2,224,447 · B 2,113,985 · C 2,158,340.

| Scenario | Post-D5 | D5 Δ (vs interconnector) | Total WS-3 Δ (vs pre-WS-3) |
|---|--:|--:|--:|
| A_Calibrated_BAU | 2,222,829 | **−6,316 (−0.283%)** | −1,618 (−0.073%) |
| B_Optimised_VRE | 2,118,643 | **+783 (+0.037%)** | +4,658 (+0.220%) |
| C_Target_VRE | 2,164,880 | **+1,753 (+0.081%)** | +6,540 (+0.303%) |

**Reading:** D5 mainly removes a placeholder artifact — the flat 5 GW badly under-stated existing internal capacity in the large India nodes (India-West peak ~73 GW had only 5 GW "existing"), forcing the optimiser to build tens of GW of unnecessary new internal transmission. Correcting it (per-node residuals) lowers cost most where least renewable transmission is needed (**A_Calibrated_BAU −0.28%**), while the **RE 2× premium** raises renewable-transmission build cost as scenarios lean renewable — roughly offsetting the residual saving in **B_Optimised_VRE (+0.04%)** and slightly outweighing it in the VRE-target **C_Target_VRE (+0.08%)**. A clean gradient A < B < C, exactly as the mechanism predicts. Net WS-3 (interconnector + internal): A ≈ neutral (−0.07%), B +0.22%, C +0.30% — all modest, all feasible (base-year backstop 0).

## P2.6 Files changed (in `OSTRAM_ws3_workcopy_D5` only)

- `t1_confection/A3_process.py` — new `stage_ws3_internal_transmission` + its call after the interconnector stage.
- `t1_confection/A3_process/rules_scripts/apply_internal_transmission.py` — **new** (the late stage; self-tested).
- `t1_confection/A3_process/rules_scripts/internal_tx_residuals.csv` — **new** (frozen desk-checked per-node residuals).
- `t1_confection/Config_country_codes.yaml` — `internal_transmission` knob block; family `OperationalLife` 20→40.

## P2.7 Still open at that checkpoint (post-D5)

- **`C_Target_VRE`:** ✔ solved (D5, optimal, feasible); delta in P2.5. **D5 is complete.**
- **D7 citations** — ✔ NP↔BD / BT↔BD backfilled in REFS `IEEE Reference` (2026-07-09; CapEx + FOM rows; 0 pending left). Optional polish left: a `Source` column in the live v18 `Interconnector_Params`, and REFS rows for the 3 added corridors.
- **Promotion:** this recorded next step was subsequently completed and is
  superseded by the accepted correction and manifest identified in the banner.

## P2.8 Methodology, verification & honest caveats

**The differential is real (not inert).** If the existing-grid residual were sized off end-year (2050) demand it would exceed demand → ~no new build → the RE/non-RE split would be inert. Checked directly: residuals use **2023** peak × 1.2 (`compute_internal_tx_residuals.py`, `YEAR=2023`); 2023→2050 demand grows **3.3–4.1×**; the solved model builds **660–930 GW of new internal transmission** per scenario, RE-weighted as expected (A 207 RE / 451 non-RE; B **615** / 311; C 445 / 415 GW). The 2× RE premium therefore bites on real build, and the cost gradient (A −0.28% / B +0.04% / C +0.08%) is the differential working.

**Uniform cost — no per-node cost premiums.** Per Luis, intra-node transmission is accounting; costs are uniform ($100 non-RE / $200 RE). The per-node cost multipliers the research produced (NP 1.8× / BT 1.6–2× / BD 1.4× / LK 1.2×, reused from generation) were **dropped** — documented, not applied. Per-node variation enters only via the (physical, computed) residuals, never the costs.

**D5 is not the ~$44 B gap.** B_Optimised_VRE vs C_Target_VRE objective gap = **44,355 pre-WS-3 → 46,237 post-D5** (WS-3 widens it ~1.9 GUSD). The internal-transmission differential (~1–2 GUSD/scenario) is an order of magnitude below the gap — internal transmission does **not** explain the B/C divergence (that is the VRE build-out the target forces).

**Grounding / confidence.** Submarine: LK↔IN 1250 Med-High (CEB LTGEP 2023, Apr-2025 MoU, Crete–Attica analog); MV↔IN 2800 **Low** (no project data; Tyrrhenian / EuroAsia / Arabian-Sea proxies) — flagged as a proxy in the v18 `Source` column + REFS. Internal $100/$200: per-kW basis (LBNL 2019: gas $44 / wind $70 / solar $103), solar ~2× — NOT the 4.5–7× per-MWh figure. RE FixedCost is **proportional** ($8 = 2× the $4 base), not flat.

**WS-1 sensitivity axes — both live.** `internal_transmission.base_capital_cost` (level; e.g. 100→185) AND `internal_transmission.re_capex_multiplier` (premium; 1.5/1.8/2/3×) are both read by the stage and can be swept independently (`base_fixed_cost` likewise).

**Dead-key landmine (recorded).** `TotalAnnualMaxCapacityInvestment` in the family blocks is NOT compiled (A2's `PARAM_LIST` omits it); a warning comment now sits in `Config_country_codes.yaml` so nobody renames it and caps transmission at 5 GW/yr.

**Provenance split.** Interconnector costs are sourced in v18 `Interconnector_Params` (+ `Source` column); internal-tx costs/knobs live in `Config_country_codes.yaml` (`internal_transmission`). "v18 = source of truth" is interconnector-scoped; the internal-tx source of truth is the YAML block. (Also in promotion handoff §3.)
