# OSTRAM — Methodology & Documentation
### Transmission calibration, model-quality fixes, and the sensitivity programme (WS-3 · WS-4 · Phase-B)

**Model.** OSTRAM is a least-cost power-system optimisation of South Asia built in OSeMOSYS, solved
with CPLEX 22.1.2. It covers **10 nodes** — Bangladesh (BGDXX), Bhutan (BTNXX), India in five regions
(INDEA/INDNE/INDNO/INDSO/INDWE), Sri Lanka (LKAXX), Maldives (MDVXX), Nepal (NPLXX) — over **2023–2050**,
with cross-border interconnectors and intra-node ("internal") transmission.

**Three baseline pathways** bracket the region's electricity future:

| Pathway | Definition |
|---|---|
| **A_Calibrated_BAU** | Business-as-usual, calibrated to observed capacity/retirement trends. |
| **B_Optimised_VRE** | Unconstrained least-cost investment — the cost-optimal frontier. |
| **C_Target_VRE** | Least-cost subject to national NDC / renewable-target generation floors. |

**Cost metric.** System cost = the sum of `TotalDiscountedCost` (M USD, NPV) over the model. This is
**not** the raw CPLEX `.sol` objective, which excludes a constant salvage/discounting term (~157,000 M USD
pre-loss; ~165,245 M USD after the transmission-loss change). All anchors in this document are on the
Σ-`TotalDiscountedCost` basis; never compare a raw `.sol` objective against them.

**Scope of this document.** The calibration and sensitivity programme layered on those baselines, in three
workstreams: **WS-3** (transmission cost/parameter calibration), **WS-4** (model-quality fixes), and
**Phase-B** (the sensitivity study), plus a literature-grounded solar-cost stress. It records the *what*,
the *why*, the *mechanism*, and the *verification*, so the work is reproducible and report-ready.

**Status note (read first).** The **foundation** (WS-3 + WS-4) is implemented, solved, and verified — its
three baseline anchors are final. The **Phase-B sensitivity results** in §4–§7 were computed on the *pre-WS-3*
foundation. The combined **15-scenario re-solve on the WS-3/WS-4 foundation is now done — 15/15 CPLEX-optimal,
zero base-year backstop** — with new anchors, the base-year-pin infeasibility fix, and the WACC test in **§8-C**.
As expected, the **absolute numbers shift up ~+100k while the behavioural story (signs & ranking) holds**; read the
§4–§7 magnitudes as pre-WS-3 and use §8-C for the current foundation.

---

## 1. Transmission calibration (WS-3)

Two long-standing weaknesses in how transmission was represented were corrected. Both are implemented as
**late core A3 stages** — they run *after* the scenario rules chain (stage 5), so they survive the earlier
template merges that would otherwise overwrite them, and they leave the post-A2 snapshot pristine.

### 1.1 Interconnector costs sourced from v18
Previously, cross-border corridor costs were legacy distance-computed values that were never actually
consumed. WS-3 makes the v18 workbook's `Interconnector_Params` sheet the **single source of truth** for
each corridor's `CapitalCost`, `FixedCost` (FOM), and `OperationalLife`:
- **18 corridors** sourced and cited (primary sources: PPAs, commissioned assets, official plans).
- Submarine links repriced (e.g. India↔Sri Lanka 1,250; India↔Maldives 2,800 USD/kW-equivalent basis).
- `OperationalLife` = **40 years**; FOM = **1.5% × CapEx**.
- A documentation `Source` column (inert to the pipeline) records the provenance of all 18 corridors in the
  live template, so the artefact self-documents.
- Stage: `apply_interconnector_costs.py`.

### 1.2 Internal (intra-node) transmission — the D5 fix
Intra-node transmission had been represented by a flat placeholder. WS-3 replaces it with a calibrated
per-node representation for the six internal families (`RNWTRN/RNWNLI/RNWRPO`, `PWRTRN/TRNNLI/TRNRPO`):
- **Per-node `ResidualCapacity`** = existing grid, sized at **peak demand × 1.2** (replaces the flat 5 GW).
  Frozen, desk-checked values in `internal_tx_residuals.csv`.
- A **2× renewable-connection CapEx premium** (basis: LBNL 2019 spur-line evidence), exposed as a live YAML
  slider so it can be varied.
- `OperationalLife` = **40 years**.
- Stage: `apply_internal_transmission.py`; knobs in `Config_country_codes.yaml → internal_transmission`.

### 1.3 Effect on the baselines
| Pathway | Pre-WS-3 | +Interconnector | +D5 (WS-3 final) |
|---|--:|--:|--:|
| A_Calibrated_BAU | 2,224,447 | 2,229,145 | **2,222,829** |
| B_Optimised_VRE | 2,113,985 | 2,117,860 | **2,118,643** |
| C_Target_VRE | 2,158,340 | 2,163,127 | **2,164,880** |

The interconnector sourcing raises costs modestly (better-costed corridors); the D5 residual/premium
calibration partly offsets it. Net effect on system cost is small (<0.3%), but the *representation* is now
physically defensible rather than placeholder-driven.

---

## 2. Model-quality calibration (WS-4)

### 2.1 Internal-transmission losses (3%)
Interconnectors already carried per-corridor losses (OutputActivityRatio 0.93–0.98), but internal
transmission carried **0%** — unphysical. WS-4 applies a **3% loss** (OAR 1.0 → **0.97**) to the six
internal families, via the same activity-ratio channel that carries interconnector losses (CEA reports
transmission losses ~3–4%; distribution is out of scope). Knob:
`Config_country_codes.yaml → internal_transmission.transmission_loss: 0.03`; stage `apply_internal_tx_losses.py`.
This is a **high-leverage** parameter — it taxes *all* throughput — and alone raised A's system cost by
**+4.1% (~$91B)**, dwarfing the interconnector/D5 effects. It is retained as a parameterised knob (a
candidate sensitivity axis).

### 2.2 Base-year lock (2023–2026 identical across scenarios)
The three pathways diverged slightly in 2023–2026 because their policy caps bite from year 1; but the near
term is calibrated history and should be identical across scenarios, with divergence only from 2027. WS-4
pins every scenario's 2023–2026 to the calibrated `A_Calibrated_BAU` solve:
- Pin **both** dispatch and build — `TotalTechnologyAnnualActivityLower/UpperLimit` **and**
  `TotalAnnualMin/MaxCapacityInvestment` — as **±0.2% bands** (exact-equality pins put CPLEX on a numerical
  knife-edge; a tight band gives clean feasibility while holding the base years identical to within 0.2%).
- Applied to **generation + mining + storage only**. **All transmission is excluded** — interconnectors are
  frozen to residual (already identical across scenarios) and internal transmission is set identically by
  D5, so they need no pin; forcing build on them collided with the interconnector residual invariant
  (`MaxCapacity ≥ Residual + Σ MinCapInvest`) and broke the pre-solve check. Excluding transmission cleared it.
- Tool: `apply_base_year_pin.py` (references the solved A-with-loss run; gated to the ~384 real technologies).
- **Verified:** all three scenarios pass GLPK `--check`, solve CPLEX-optimal, base-year backstop = 0, and
  2023–2026 generation *and* capacity are identical across A/B/C within the band.

### 2.3 C_Target build-ahead cliff relax
On C_Target, two near-term VRE floors exceeded what the model could physically build in time
(`PWRWONINDWE`, 2027 and 2028 — the NDC floor sat above the buildable ramp given residual + investment
limits). The floors were lowered to **0.98 × the buildable maximum** (2027: 237.221→**227.414**; 2028:
260.745→**255.078**), cross-checked with the CPLEX conflict refiner. This clears the infeasibility while
keeping the floors non-binding at essentially their physical ceiling (the 2% headroom cost zero realised ambition).

### 2.4 Foundation anchors (WS-3 + WS-4, final — validated)
| Pathway | Final system cost (Σ TotalDiscountedCost, M USD) |
|---|--:|
| A_Calibrated_BAU | **2,314,332** |
| B_Optimised_VRE | **2,215,073** |
| C_Target_VRE | **2,257,995** |

Cost ranking A > C > B holds: BAU is dearest, the VRE-optimised pathway cheapest, the target pathway in
between. Caveat: base-year "identical" means within the 0.2% band, not byte-exact; and internal-transmission
per-corridor routing is degenerate (annual totals match; individual lines may differ).

---

## 3. Physical-potential VRE ceiling (the "clips")

A shared ceiling is layered on every scenario: `TotalAnnualMaxCapacity(node,tech) = min(atlas_potential,
B_Opt MaxCap)`, flat across years, for solar (`PWRSPV`) and onshore wind (`PWRWON`); atlas from NISE-2025
(solar) and NIWE-150 m (wind). Taking the *minimum* is a **pure guard** — it only trims overbuild, never
relaxes a bound, and preserves B_Opt's own zeros. Three nodes are clipped where B_Opt overbuilt past the
atlas and the atlas is enforced: **PWRSPVLKAXX 16, PWRWONBGDXX 3, PWRWONMDVXX 0**. Where a clip lowers a
MaxCap below an NDC generation floor's envelope (only Maldives solar), the activity *lower* limit is scaled
down by the same ratio so the target stays physically buildable (upper limits are never scaled).

**Cost of physical realism (clipped − unclipped, pre-WS-3 basis):** A +0.00%, B +0.05%, C −0.43%. Enforcing
real limits is essentially free on B and *cheaper* on the target pathway (it removes unbuildable obligations).

## 4. Sensitivity programme (Phase-B)

Sensitivities branch one-at-a-time from the ceiling-only baseline (`B_Opt_Clipped`), isolating each lever
from the clip. The centrepiece is the **energy-security question**: the cost-optimal pathway leaves
**Bangladesh importing ~62% of its 2050 electricity** — far above today's ~10% and its own ~15–20% plans.

**The three interconnection levers:**
| Lever | Mechanism (one thing changes) |
|---|---|
| **Volume** — `B_Opt_TradeCap15` | Bangladesh imports ≤ 15% of demand (strict; no export allowance); backstops disabled |
| **Capacity** — `B_Opt_TxCap150` | every cross-border corridor MaxCap = 1.5 × 2023 residual; India-internal kept; non-India backstops disabled |
| **Direction** — `B_Opt_DirContractual` | each corridor forced to its real contractual one-way flow (9 one-way, 2 Nepal↔India bidirectional) |

Direction is set from **primary sources per corridor** (Adani Godda PPA, Bheramara HVDC, Bhutan hydro export,
the 2024 Nepal–Bangladesh trilateral, etc.); conceptual corridors get the only physically plausible
(deficit-import) direction. A neutrality scenario (`B_Opt_DirBidir`, empty map) reproduces `B_Opt_Clipped`
to the decimal, proving the direction machinery is itself neutral.

**Cost / robustness sensitivities:** `B_Opt_SolarCapexHi` (solar CapEx ×1.10), `B_Opt_IndiaCosts` (non-India
generation/storage capital+fixed costs set to India's reference), `B_Opt_IndiaCostsFuel` (+ India fuel prices).

**Findings (pre-WS-3 solve; behavioural, to be re-confirmed on the new foundation):**
- Quantity levers buy security cheaply: the volume cap (+0.5%) and capacity cap (+1.3%) pull Bangladesh from
  38% to **87%–95% self-sufficient** with zero unserved energy, and *lower* CO₂ (self-supply displaces
  coal-heavy imports; the volume cap is the lowest-emission scenario).
- The direction correction runs the other way: honouring real one-way contracts **deepens** dependence
  (→32% domestic) — part of B_Opt's "domestic" build was really build-for-export.
- The result is **robust** to a +10% solar shock (mix barely moves); the biggest single swing is the
  **India-vs-neighbour cost gap** (−1.8%), implying some of Bangladesh's import reliance is a cost-assumption
  artefact. `IndiaCostsFuel` reproduces `IndiaCosts` exactly (fuel-price harmonisation non-binding).

Full detail: `t1_confection/sensitivity_expansion/PHASE_B_METHODOLOGY_AND_RESULTS.md` and
`docs/archive/phase-b/PHASE_B_IMPLEMENTATION_LOG.md`.

---

## 5. Solar CapEx stress (literature-grounded)

The ×1.10 solar shock proved immaterial (smaller than one year of normal module deflation). A dedicated
literature review established two defensible adverse tiers on utility-scale solar PV **overnight CapEx**:
- **Sustained / structural: ×1.30** (range ×1.20–1.40) — the primary stress. Independently corroborated by
  the NREL ATB Conservative-vs-Moderate CapEx spread (~30% at both 2035 and 2050), the IEA supply-chain
  diversification premium (India ≈ +11%, up to ~25%), and a WACC-stress equivalent (5%→9% ≈ +34% on LCOE).
- **Transient / spike: ×1.50** (range ×1.40–1.70) applied to a **1–3-year window (2028–2030) with reversion**
  — a stacked polysilicon + trade-measure + materials shock, as observed in 2021–2022.

Modelling notes: the multiplier rides on the model's existing South-Asia-calibrated per-node base CapEx (the
base is *not* changed). WACC/financing stress is handled separately as a discount-rate sensitivity, not
folded into the CapEx knob. The 2021–22 episode passed through to installed cost by only ~8% and reverted in
~18 months — so the transient tier is expected to show little cumulative-2050 effect, which is itself a
defensible robustness finding. The **breakeven CapEx multiple** (where solar stops being cheapest vs firm)
is read off the model by running the ×1.20/1.30/1.40 band and reporting where the build mix shifts.

## 6. The full scenario set (15) on the new foundation

| Tier | Scenarios |
|---|---|
| Baselines (3) | A_Calibrated_BAU, B_Optimised_VRE, C_Target_VRE |
| Clipped baselines (3) | A_Calibrated_BAU_Clipped, B_Opt_Clipped, C_Target_VRE_Clipped |
| Interconnection levers (3) | B_Opt_TradeCap15 (volume), B_Opt_TxCap150 (capacity), B_Opt_DirContractual (direction) |
| Direction control (1) | B_Opt_DirBidir (neutrality check) |
| Cost/robustness (3) | B_Opt_IndiaCosts, B_Opt_IndiaCostsFuel, B_Opt_SolarCapexHi (×1.10) |
| Solar stress (2) | B_Opt_SolarCapex130 (×1.30 sustained), B_Opt_SolarCapexSpike (×1.50 transient) |

All 15 inherit the WS-3+WS-4 foundation (transmission costs, losses, base-year pin) and the VRE ceiling.

## 7. Reproducibility & verification

- **Anchor definition:** Σ `Outputs/TotalDiscountedCost.csv` (not the raw `.sol`; constant offset ~165,245 post-loss).
- **No-solve reproduction evidence:** compile a scenario's datafile without solving (`execute_model:False`),
  normalize supported line-ending differences, and compare it with the corresponding historical input.
  A zero diff establishes exact pre-solver input equivalence; it does not guarantee numerical identity
  across solver versions or replace a CPLEX-backed behavioral baseline.
- **Hero repos (read-only oracles):** `OSTRAM_latest` (pre-WS-3), `OSTRAM_ws3_workcopy_D5` (WS-3),
  `OSTRAM_ws4_workcopy` (WS-4 final + cumulative source).
- **Clean-room rebuild:** the canonical repo is reconstructed from source in `OSTRAM_mainredo`, verified by
  byte-diff at each layer, with a self-checking test harness and per-step checkpoint commits
  (see `docs/archive/cleanroom/CLEANROOM_FINALPROMPT.md`). The CPLEX solves are staged as a batch, not run inline.
- **Solve settings:** CPLEX 22.1.2, StorageDelayN5 variant, `cplex_threads: 4`; every reported solve is
  CPLEX-optimal with zero base-year backstop generation.

---

## 8. Gaps & what to bring for an expanded report

### A. Must resolve before publishing final sensitivity numbers
1. **Combined 15-scenario re-solve on the WS-3/WS-4 foundation is pending.** The Phase-B magnitudes in §4
   are from the *pre-WS-3* solve. The report's headline sensitivity numbers must come from the clean-redo
   batch solve (staged in `docs/archive/cleanroom/CLEANROOM_FINALPROMPT.md`). Expect the *story* to hold, the *numbers* to move.
2. **Unclipped-B_Opt Capital(NPV) concat glitch.** That one cell reads ~6.89M (≈10× the others); it is an
   aggregation artefact and must be fixed before any capital-vs-opex breakdown is published. System-cost
   anchors are unaffected (benchmarked against clean values).

### B. Base-model documentation to source from upstream (not covered here)
This document covers the *calibration + sensitivity* layer. A full methodology report also needs the base
model's own documentation, which lives in the v18 template / upstream OSTRAM docs, not here:
3. **Demand projections** (per-node electricity demand trajectory 2023–2050 and its basis).
4. **Discount rate(s)** and the salvage/depreciation treatment (state the value; ties to the WACC point).
5. **Temporal resolution** — timeslice structure (season × daytype × daily-time-bracket), storage/StorageDelay representation, reserve margin.
6. **Technology set & techno-economic assumptions** — the base CapEx/FOM/VOM/efficiency/lifetime per family and node (especially the base solar $/kW the ×1.30/×1.50 multipliers ride on).
7. **Original A/B/C calibration methodology** — how the BAU was calibrated and how the NDC/RE targets in C were sourced and encoded.
8. **Cost base year & real-vs-nominal convention** — state the USD base year and deflation treatment (the solar review flagged real-vs-nominal care explicitly).
9. **Emissions accounting** — emission factors and whether any carbon price/constraint is applied (results cite CO₂).

### C. Would strengthen the report
10. **WACC / discount-rate sensitivity** — flagged but not run; high-value given the India-cost finding. A ±300–400 bps sweep would complement the solar-CapEx stress.
11. **External-benchmark validation** — compare baseline capacity/generation to national plans (India CEA/LTGEP, Bangladesh IEPMP, etc.) to evidence credibility.
12. **Figures** — an anchor "waterfall" (pre-WS3 → +interconnector → +D5 → +loss → final), a Bangladesh self-sufficiency-vs-cost scatter across the levers, capacity-mix stacked bars, and a corridor-flow/direction map. I can generate these on request.
13. **Market-specific cost uplifts** — smaller markets (BGD/NPL/LKA/BTN/MDV) have thin published CapEx data (flagged by WS-3 and the solar review); consider node-specific uplifts and state the caveat.

### D. Known limitations to state honestly
14. **Base-year "identical" = within the ±0.2% pin band**, not byte-exact.
15. **Internal-transmission routing is degenerate** — annual totals match across scenarios; individual corridor flows may differ.
16. **TxCap150 blocks zero-residual corridors** (BTN→BGD, IND-S→LKA) that B_Opt used; a softer "pipeline-floor" variant is noted but not built.
17. **Two Maldives corridors are conceptual** — directions inferred, non-binding (no committed project).
18. **Node-mapping** — Dhalkebar–Muzaffarpur physically lands in India's Eastern region; a mapping review is open (non-binding, both bidirectional).
19. **Transient solar spike** needs a per-year CapEx multiply (a small patcher extension) — not yet built.

---

## 8-C. Clean-redo results — WS-3/WS-4 re-solve, the 2027+ infeasibility fix, and the WACC test

*(Session 2, 2026-07-12, branch `ws3-phaseb-cleanredo`. This section supersedes the "pending re-solve"
caveat in §8-A.1 and the "WACC flagged but not run" note in §8-C.10 above: the combined 15-scenario solve on
the WS-3/WS-4 foundation is now **done — 15/15 CPLEX-optimal, zero base-year backstop generation** — and the
WACC mechanism is proven.)*

### 8-C.1 Method
All 15 scenarios were rebuilt on the WS-3/WS-4 foundation (`apply_base_year_pin.py --band 0.002` baked into the
pin) and solved with CPLEX 22.1.2, **dual simplex, feasopt-off** (barrier was tried and reverted — 3–4× slower on
this LP). System cost = Σ`TotalDiscountedCost`. Sensitivities branch from the pinned B_Optimised_VRE A-O via
`apply_patches.py`; the Dir scenarios overlay `set_interconnector_direction.py`. Each scenario compiled with
`glpsol --check` clean before solving; every solve is Optimal with **backstop generation = 0**.

### 8-C.2 The base-year-pin × network-lever infeasibility, and the fix
Three sensitivities restrict **base-year network flows** — `TradeCap15` (import-volume cap), `TxCap150`
(corridor-capacity cap), `DirContractual` (one-way corridor directions). On the new foundation these were
initially **infeasible**, each on a base-year power-backstop upper limit (`PWRBCK{BGDXX,BGDXX,BTNXX}` @
2024/2025/2023). The cause is a genuine pin×lever conflict, not a pipeline bug: the WS-4 base-year pin freezes
each node's **calibrated 2023–2026 mix**, which fixes its base-year *net import* (e.g. Bangladesh needs ~242 PJ of
imports in 2024 = demand 456 − pinned domestic 214). A lever that cuts base-year flows then leaves that demand
unmeetable, and the emergency power backstop is pinned to 0 → infeasible. (These solved in the pre-WS-3 Phase-B
because there was no base-year pin.)

**Fix (modeling call): apply each lever to the STUDY PERIOD 2027+ only; leave 2023–2026 = the pinned calibrated mix.**
- `TradeCap15` (per-year *activity* cap) and `DirContractual` (per-year *activity-ratio* direction) are non-cumulative,
  so the fix is a clean 2027+ restriction: corridor `TotalTechnologyAnnualActivityUpperLimit` values omit 2023–2026
  (revert to the source `−1` = unconstrained), backstop-import zeros start 2027, and `set_interconnector_direction.py`
  gains `--study-start-year 2027` (zeros the disabled mode only for years ≥ 2027; the base-year AR file is left
  bidirectional).
- `TxCap150` caps `TotalAnnualMaxCapacity`, a **cumulative, long-lived** variable (interconnector life 40 y). Because
  base-year builds persist, a bare 2027 cap retroactively starves the pinned base years, and 1.5×Residual₂₀₂₃ (~4 GW
  across Bangladesh's corridors) physically cannot carry the pinned ~242 PJ base-year import (~8 GW). The faithful
  realisation of "leave base years calibrated" is therefore to **grandfather** the cap: for 2027+,
  `MaxCap = max(1.5×Residual₂₀₂₃, calibrated base-window capacity)`. Only Bangladesh's two over-built corridors are
  grandfathered (`TRNBGDXXINDEA` 3.75→8.24 GW, `TRNBGDXXINDNE` 0.24→0.32); all others keep 1.5×Residual. This freezes
  study-period expansion (B_Opt grows `TRNBGDXXINDEA` to 25.4 GW; here it holds 8.24) without breaking the base window.
  *Consequence to note in the report: TxCap150's capacity cut on Bangladesh's main corridor is softer than the
  pre-WS-3 oracle (which had no base-year pin); the direction and ranking are unchanged, the magnitude is smaller.*

### 8-C.3 Anchors (WS-3/WS-4 foundation, Σ-TotalDiscountedCost, M USD)
| Scenario | Σ-TDC | Δ vs Clipped | (pre-WS-3 oracle Δ) |
|---|--:|--:|--:|
| A_Calibrated_BAU | 2,314,128 | +98,133 | (BAU, dearest) |
| A_Calibrated_BAU_Clipped | 2,314,131 | +98,136 | — |
| B_Optimised_VRE | 2,214,920 | −1,076 | — |
| **B_Opt_Clipped** *(baseline)* | **2,215,995** | **0** | 0 |
| C_Target_VRE | 2,257,930 | +41,934 | (+ve) |
| C_Target_VRE_Clipped | 2,246,158 | +30,163 | (+ve) |
| B_Opt_TradeCap15 | 2,224,144 | +8,148 (+0.37%) | +0.5% |
| B_Opt_TxCap150 | 2,239,553 | +23,558 (+1.06%) | +1.3% |
| B_Opt_DirContractual | 2,224,494 | +8,498 (+0.38%) | +0.2% |
| B_Opt_SolarCapexHi | 2,223,239 | +7,243 | +0.3% |
| B_Opt_SolarCapex130 | 2,237,715 | +21,720 | *(new)* |
| B_Opt_SolarCapexSpike | 2,220,882 | +4,886 | *(new)* |
| B_Opt_IndiaCosts | 2,177,458 | −38,537 | −1.8% |
| B_Opt_IndiaCostsFuel | 2,177,458 | −38,537 | −1.8% |
| B_Opt_DirBidir *(validation)* | 2,215,995 | 0 | 0 |

*All absolute anchors are ~+100k above the pre-WS-3 values — the WS-3/WS-4 foundation (3% transmission loss +
interconnector cost/parameter calibration) shifts the whole frontier up; the deltas and behaviour are what carry over.*

### 8-C.4 Behavioural cross-check vs the pre-WS-3 Phase-B oracle — signs & ranking hold
- **Bangladesh self-sufficiency (2050) direction — all match:** TradeCap15 **↑**, TxCap150 **↑**, `DirContractual` **↓**
  (still the counter-intuitive result — honouring real one-way contracts removes build-for-export and *deepens*
  dependence), IndiaCosts **↑**, DirBidir **~same**.
- **Neutralities hold to the dollar:** `IndiaCostsFuel ≡ IndiaCosts` (2,177,458; non-India fuel differentials
  non-binding) and `DirBidir ≡ B_Opt_Clipped` (2,215,995; direction machinery neutral).
- **Solar-cost tiers order correctly:** transient Spike (×1.5 in 2028–30, +4,886) < sustained Hi (×1.10, +7,243) <
  sustained 130 (×1.30, +21,720) — sustained > transient, and cost rises with the multiplier.
- **CO₂ (2050) order matches:** BAU (3,859 Mt) ≫ C_Target (2,685) > B-family (~1,890), with TradeCap15 the lowest of
  all (1,743 Mt; self-supply displaces coal-heavy imports).
- **Coarse system-cost ranking matches:** IndiaCosts (cheapest) → B_Opt cluster → interconnection/solar sensitivities
  → C_Target → BAU (dearest). The only reshuffle is *within* the tight ±0.3–0.5% sensitivity cluster: DirContractual
  moved from the pre-WS-3 cheapest lever (+0.2%) to ≈TradeCap15 (+0.38% vs +0.37%), because WS-3/WS-4 raised
  transmission costs that the direction lever interacts with. Sign unchanged; a sub-0.1% within-cluster move.

Reproduce with `t1_confection/sensitivity_expansion/analyse_ws4_vs_phaseB.py` (reads the
solved `t1_confection/Executables/<s>_0/Outputs/`). This analysis-only script remains in
place because the protected-tree gate covers its current path.

### 8-C.5 WACC / discount-rate test (mechanism proof) — ✅ PASS
`B_Opt_Clipped` re-solved with DiscountRate + DiscountRateStorage **0.10 → 0.13** (a single knob; injected via the
otoole `DiscountRate.csv`/`DiscountRateStorage.csv` with `A2_otoole_outputs:False` so B2 consumes the edit rather than
regenerating the header-only template; verified `DiscountRate := GLOBAL 0.13` in the compiled `.txt`; CPLEX Optimal,
backstop 0).

- **Σ-TotalDiscountedCost: 2,215,995 (10%) → 1,761,993 (13%), Δ −454,002 (−20.5%).** The knob is unambiguously live —
  a higher discount rate weights 2027–2050 costs less (a 2050 cost is valued ~half as much at 13% as at 10%), so the
  NPV of total system cost falls.
- **2050 build-mix shift is directionally correct but small:** solar 878.3→874.0 GW (−4.3), oil +0.5, hydro +0.6,
  CO₂ +7.6 Mt — VRE edges down, firm/fossil edges up. The small magnitude is itself the result: the ceiling-clipped
  VRE stays far cheaper than firm even at 13%, so the least-cost pathway's reliance on cheap VRE is **not fragile** to
  a +300 bps WACC shock (the cost-of-capital restatement of the §5.4 robustness finding). The same mechanism extends to
  a 7%/13% × {B_Opt_Clipped, TradeCap15, TxCap150, DirContractual} matrix; full numbers in
  `docs/archive/validation/WACC_TEST_RESULT.md`.

---

## 9. Source files (where each fact lives)
- **This programme:** `docs/archive/phase-b/PHASE_B_IMPLEMENTATION_LOG.md`, `t1_confection/sensitivity_expansion/PHASE_B_METHODOLOGY_AND_RESULTS.md`; `docs/archive/ws3-ws4/WS3_PROMOTION_HANDOFF.md`, `ws3_transmission_audit/WS3_calibration_report.md`, `ws3_transmission_audit/WS3_value_audit.md`, `docs/archive/ws3-ws4/WS4_HANDOVER_PROMPT.md`, `docs/archive/ws3-ws4/WS4_PREFLIGHT.md`; `t1_confection/sensitivity_expansion/reference/interconnector_direction_references.md` (cited corridor directions).
- **VRE-ceiling inputs and provenance:** `t1_confection/sensitivity_expansion/reference/vre_ceilings_base.json` is the operational source of truth consumed by `apply_patches.py`. `t1_confection/sensitivity_expansion/reference/vre_ceilings.csv` is a documentary mirror of the same 20 ceiling values, with B_Opt, clip, atlas/source-label, confidence, and metadata columns. Read the CSV together with `t1_confection/sensitivity_expansion/reference/vre_ceiling_provenance.md`, which records source support, caveats, the India-only validity of the NISE/NIWE labels, and remaining citation gaps. A root `reference/` directory and root `reference/vre_ceilings.csv` never existed in available Git history; the old methodology wording was stale shorthand, not evidence of a former root path. Generated workbooks, compiled CSVs, change manifests, validation reports, and sensitivity tables are derived evidence, not primary provenance.
- **Rebuild recipe:** `docs/archive/cleanroom/CLEANROOM_FINALPROMPT.md`.
- **Pre-WS-3 analysis outputs:** `t1_confection/sensitivity_report.txt` and `t1_confection/sensitivity_comparison.csv`. Narrative tables are also embedded in `t1_confection/sensitivity_expansion/PHASE_B_METHODOLOGY_AND_RESULTS.md`.
- **Solved anchors:** hero repos `OSTRAM_latest`, `OSTRAM_ws3_workcopy_D5`, `OSTRAM_ws4_workcopy`.

*Prepared as report-feeding methodology documentation. Foundation anchors are validated; the §4–§7 sensitivity
magnitudes are pre-WS-3, and the WS-3/WS-4 re-solve (15/15 optimal), the base-year-pin infeasibility fix, and the
WACC test are confirmed in §8-C.*

## 10. Provenance & citations update (2026-07-12 research pass)

A verification pass sourced the study's previously-uncited, load-bearing assumptions. **No ceiling or model
value was changed** — this records provenance and corrects labels/wording only. Full reference list + per-node
support table: `t1_confection/sensitivity_expansion/reference/vre_ceiling_provenance.md`. Key items:

- **VRE ceilings — label fix:** `t1_confection/sensitivity_expansion/reference/vre_ceilings.csv` tags all solar "NISE-2025" / all wind "NIWE 150 m", but NISE/NIWE are **India-only** (MNRE). India stands (solar → NISE 2025, 3,343 GWp; wind → NIWE 150 m, 1,163.9 GW — our India totals reconcile). The **five non-India nodes** need their own sources — cleanest is a **Global Solar/Wind Atlas** (World Bank/ESMAP·DTU) baseline + a national study on each clip.
- **Clip caveats (state in the report; none changes a result):** LKA solar 16 GW is a 2050 *scenario-build* figure (ADB/UNDP), not pure technical potential (~6 GW); **BGD onshore wind 3 GW is a modeler's conservative screen — a cite-gap** (only a gross ">30 GW, unrealistic" figure exists); **MDV onshore wind ≈0** is defensible but not literally zero (~80 MW niche, Greater Malé).
- **Transmission loss:** recast "~3%" → "**~3–4%**" per Grid-India ISTS notices (3.96% recent), and **drop "single largest driver of system cost"** (over-claim — generation capital+fuel dominate). §2.1's model result "+4.1% on A, dwarfs the interconnector/D5 effects" is correctly scoped and stands.
- **WACC 10%:** defensible via IRENA "10% rest-of-world" (RPGC 2018); conservative/high end of the ~4–9% (OECD/IEA) to 6% (World Bank) range; no South-Asia-specific single source.
- **D4 refs:** [27] (Koirala & Rahut 2022) is an **ADBI blog, not peer-reviewed** — re-tag as grey literature; [25]/[26] are the journal + working-paper versions of one study (add middle initials to [26]).

See `t1_confection/sensitivity_expansion/reference/vre_ceiling_provenance.md` for full citations, URLs, and the per-node support table.
