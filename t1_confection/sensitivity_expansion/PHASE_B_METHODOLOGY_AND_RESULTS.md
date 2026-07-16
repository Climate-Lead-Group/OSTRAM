<!-- CLG / OSTRAM -->
# OSTRAM Phase-B Sensitivity Analysis — Methodology, Results & Implications

**CLG · OSTRAM** — 10-node OSeMOSYS South Asia model (BGD, BTN, IND×5, LKA, MDV, NPL).
Comprehensive record of the Phase-B sensitivity set: what was run, how, what it shows, and
why each lever matters. Companion to the build-side paper trail in
[`PHASE_B_IMPLEMENTATION_LOG.md`](../../docs/archive/phase-b/PHASE_B_IMPLEMENTATION_LOG.md).

> **Status:** 13 scenarios, all solved optimal (CPLEX 22.1.2), zero backstop generation, all
> validated. Results drawn from `OSTRAM_StorageDelay_Combined_Inputs_Outputs_2026-07-09.csv`
> via `analyse_sensitivity.py` (`sensitivity_report.txt` / `sensitivity_comparison.csv`).
> System cost = Σ `TotalDiscountedCost` (M USD, NPV) = CPLEX objective + 157,222 constant term.

---

## Contents
1. [Executive summary](#1-executive-summary)
2. [Background: the question the sensitivities answer](#2-background-the-question-the-sensitivities-answer)
3. [Scenario architecture](#3-scenario-architecture)
4. [Methodology](#4-methodology)
5. [Results](#5-results)
6. [Implications](#6-implications)
7. [Relevance: why each sensitivity was chosen](#7-relevance-why-each-sensitivity-was-chosen)
8. [Caveats & limitations](#8-caveats--limitations)
9. [Reproducibility](#9-reproducibility)
- [Appendix A — Full metric matrix (13 scenarios)](#appendix-a--full-metric-matrix-13-scenarios)
- [Appendix B — Bangladesh 2050 capacity by scenario](#appendix-b--bangladesh-2050-capacity-by-scenario)
- [Appendix C — Interconnector direction map](#appendix-c--interconnector-direction-map)

---

## 1. Executive summary

The validated cost-optimal decarbonisation pathway for South Asia (**B_Optimised_VRE**) leaves
**Bangladesh importing ~62 % of its 2050 electricity** — far above today's ~10 % and above even
Bangladesh's own ~15–20 % import plans. Phase-B stress-tests that result along the dimension that
matters most for a national planner — **cross-border dependence** — plus the cost assumptions that
the VRE-dominated outcome rests on. Three findings stand out:

- **Physical VRE realism is close to free.** Clipping the model's VRE build to atlas-based
  physical potential (correcting an unphysical 9.16 GW of Bangladesh wind, a 24.9 GW Sri Lanka
  solar overbuild, etc.) costs **+0.05 %** on the B pathway and is **−0.43 % (cheaper)** on the
  NDC-target pathway. Enforcing real limits does not distort the least-cost story.
- **Energy security is cheap to buy — through quantity levers.** Capping Bangladesh's import
  *volume* at 15 % of demand (**+0.5 %** system cost) or its import *capacity* at 1.5× today's
  corridors (**+1.3 %**) pulls Bangladesh from 38 % to **87 %–95 % self-sufficient**, with **zero
  unserved energy** and **lower CO₂** (self-supply displaces coal-heavy imports; the 15 % volume
  cap is the lowest-emission scenario of all thirteen).
- **Direction and cost assumptions cut the other way.** Enforcing the *real, contractual one-way*
  flows of the region's interconnectors (**+0.2 %**) actually **deepens** Bangladesh's dependence
  (to 32 % domestic) — revealing that part of B_Opt's "domestic" generation was really
  build-for-export. And harmonising the smaller nations' capital costs down to India's
  (**−1.8 %**) lifts Bangladesh to 51 % domestic — implying its import reliance is **partly a
  cost-assumption artifact**, not pure geography.

The VRE-dominated result itself is **robust**: a +10 % solar-CapEx shock barely moves the capacity
mix (+0.3 % cost, solar unchanged). The single largest cost swing in the whole set is not a
technology cost but the **India-versus-neighbour cost gap**.

---

## 2. Background: the question the sensitivities answer

OSTRAM's three validated pathways bracket the region's electricity future to 2050:

| Pathway | Definition |
|---|---|
| **A_Calibrated_BAU** | Business-as-usual, calibrated to observed capacity/retirement trends. |
| **B_Optimised_VRE** | Unconstrained least-cost investment — the cost-optimal frontier. **Anchor.** |
| **C_Target_VRE** | Least-cost *subject to* national NDC / RE-target generation floors. |

B_Opt is the natural reference because it is the cost floor. But two features of its solution are
not physically or politically self-evident and must be tested before the pathway can be used to
advise policy:

1. **VRE overbuild beyond physical potential.** B_Opt builds VRE past what solar/wind atlases say
   a node can host (e.g. Bangladesh onshore wind at **9.16 GW** against a NIWE-150 m realisable
   **3 GW**; Sri Lanka solar at **24.9 GW** against **16 GW**). Left uncorrected, the least-cost
   frontier rests partly on capacity that cannot be built.
2. **Extreme cross-border dependence.** B_Opt meets **~62 %** of Bangladesh's 2050 demand with
   imports. For the region's only large net importer, that is an energy-security exposure a
   sovereign planner would not accept at face value — and it is **6× today's level**.

Phase-B therefore layers a **shared physical-potential ceiling** onto every scenario, then runs a
family of one-lever-at-a-time sensitivities that (a) price the cost of that realism, (b) test how
cheaply — and by which policy instrument — Bangladesh's dependence can be pulled to realistic
levels, and (c) probe the robustness of the VRE-dominated result to its key cost assumptions.

---

## 3. Scenario architecture

Thirteen scenarios in three tiers. All sensitivities branch from **B_Optimised_VRE** and are
measured against **B_Opt_Clipped** (the VRE-ceiling-only baseline), isolating each lever from the
clip effect.

### Baselines — 3 pathways × {unclipped, clipped}
| Scenario | What it is |
|---|---|
| `A_Calibrated_BAU` / `A_Calibrated_BAU_Clipped` | BAU, without / with the VRE ceiling. |
| `B_Optimised_VRE` / **`B_Opt_Clipped`** | Cost-optimal, without / with the ceiling. **Clipped = comparison baseline.** |
| `C_Target_VRE` / `C_Target_VRE_Clipped` | NDC-target-constrained, without / with the ceiling. |

### The three interconnection levers *(the centrepiece)*
| Lever | Scenario | Mechanism (one thing changes) |
|---|---|---|
| **Volume** | `B_Opt_TradeCap15` | Bangladesh imports ≤ **15 %** of demand (strict — no export allowance). |
| **Capacity** | `B_Opt_TxCap150` | Every cross-border corridor MaxCap = **1.5 × Residual₂₀₂₃**. |
| **Direction** | `B_Opt_DirContractual` | Each corridor forced to its **real contractual one-way** flow. |

### Cost / robustness sensitivities + machinery validation
| Scenario | Mechanism |
|---|---|
| `B_Opt_SolarCapexHi` | Solar CapitalCost **× 1.10**, all `PWRSPV`, all nodes/years. |
| `B_Opt_IndiaCosts` | Non-India generation/storage **CapitalCost + FixedCost → India reference**. |
| `B_Opt_IndiaCostsFuel` | As above **plus** non-India fuel-supply prices → India prices. |
| `B_Opt_DirBidir` | Direction machinery with an **empty** map — must reproduce `B_Opt_Clipped` (neutrality test). |

---

## 4. Methodology

### 4.1 Branch-and-patch design
Every sensitivity is the validated B_Opt A–O parametrisation **plus a small, declarative set of
post-A3 edits** (`patches.json`), applied by `apply_patches.py` onto a fresh copy of the B_Opt
inputs. This keeps each run one auditable step from the anchor: the diff *is* the lever. Patches
are generated deterministically by `gen_sensitivity_patches.py` (corridor lists, residuals and
India reference costs are all **read from the A-O**, nothing hardcoded), so the whole set is
reproducible from source. `apply_patches.py` is idempotent, non-destructive to the source B_Opt,
and supports `--restore` / `--self-test`.

> **Why post-A3 patches and not YAML rule edits.** The A3 `relax_interconnectors` rule has a
> second pass that lifts any overridden corridor's `TotalAnnualMaxCapacity` to a 9999 placeholder —
> so a corridor cap expressed as a YAML override is silently removed. Capacity/freeze levers are
> therefore applied as explicit post-A3 parameter edits, which the compiler cannot override.

### 4.2 Shared VRE physical-potential ceiling
A single ceiling layer is applied to **every** scenario:

> `TotalAnnualMaxCapacity(node, tech) = min(atlas_potential, B_Opt MaxCap)`, flat across all years,
> for solar (`PWRSPV`) and onshore wind (`PWRWON`).

- **Atlas sources:** NISE-2025 (solar), NIWE-150 m (onshore wind).
- **Pure guard:** taking the *min* never relaxes B_Opt — it only trims overbuild — so it cannot
  make any scenario cheaper by loosening a bound, and it preserves B_Opt's own zeros
  (`PWRWON INDEA/INDNE/NPLXX = 0`).
- **Three deliberate clips** where B_Opt overbuilds past the atlas and the atlas is enforced:
  `PWRSPVLKAXX 16` (was 24.9), `PWRWONBGDXX 3` (was 9.16), `PWRWONMDVXX 0`. These are a physical-
  realism correction to the shared base, accepted on Luis's Step-0 decision.
- **Ceiling ↔ NDC-floor coherence (pathway C only).** Where the clip lowers a MaxCap below the
  cap-envelope that an NDC generation floor was sized against, `apply_patches.py` scales the
  **activity lower limit** down by the same `ceil / MaxCap_orig` ratio (production is linear in
  capacity). Only `PWRSPVMDVXX` is affected (Maldives solar floor re-capped to the physically
  buildable 1 GW ≈ 4.22 PJ). The upper limit is **never** scaled (B_Opt's `-1` "unconstrained"
  sentinel would be corrupted).

### 4.3 Clipped baseline of comparison
All sensitivity deltas are quoted against **B_Opt_Clipped**, not raw B_Opt. This is deliberate:
every sensitivity *also* carries the ceiling, so comparing to the clipped baseline isolates the
lever from the clip. The cost of the clip itself is reported separately (§5.2).

### 4.4 The three interconnection levers (mechanisms)

**Volume — `B_Opt_TradeCap15`.** Bangladesh is the region's **only large net importer** in B_Opt
2050 (Sri Lanka, Nepal, Bhutan are net exporters; India and Maldives self-sufficient), so a
region-wide import cap binds on Bangladesh alone. The four real import corridors carry a per-year
`TotalTechnologyAnnualActivityUpperLimit` = `15 % × demand × (that corridor's B_Opt import share)`;
backstop imports (`TRNNLIBGDXX`, `TRNRPOBGDXX`) are zeroed. **Strict** means the ×1.5 export
allowance used in earlier drafts is dropped (`export_factor = 0`) — without it "15 %" is a true
≤15 % *import* cap. (With the allowance, the optimiser spent export headroom on imports and landed
at ~18 %.)

**Capacity — `B_Opt_TxCap150`.** Every cross-border corridor's `TotalAnnualMaxCapacity` is set to
**1.5 × its 2023 residual capacity**, flat across all years (with `MaxCapInv` matched and
`MinCapInv` clamped to the `MaxCap − Residual` headroom for solver coherence). India-internal
corridors are untouched (intra-national transfers ≠ imports); non-India backstops are activity-
blocked (`AUL = 0`). This is the *physical wires* form of the volume cap: you cannot import more
than the corridor can carry.

**Direction — `B_Opt_DirContractual`.** Each corridor has two flow modes; B_Opt treats every one as
freely bidirectional. `set_interconnector_direction.py` zeroes the disabled mode's
Input/Output ActivityRatio so each corridor carries only its **real governing flow**, established
from primary sources (PPAs, commissioned assets, official plans) over three research passes. **9
corridors locked one-way, 2 (Nepal↔India) left bidirectional** (genuinely seasonal). See
[Appendix C](#appendix-c--interconnector-direction-map) and
[`reference/interconnector_direction_references.md`](reference/interconnector_direction_references.md)
for the cited, per-corridor justification.

### 4.5 Cost & robustness sensitivities (mechanisms)

**`B_Opt_SolarCapexHi`** — a single knob multiplies solar `CapitalCost` by **1.10** across all ten
nodes and all years, riding on top of the per-node cost multipliers already in the A-O. It tests
the load-bearing "VRE is cheap" assumption.

**`B_Opt_IndiaCosts`** — the `CapitalCost` + `FixedCost` of every **non-India** generation/storage
family is overwritten with the **India reference trajectory** (per family and year; INDNO anchor
where the five India nodes disagree — e.g. `PWRCOA`, `PWRHYD`, `PWRWOF`). Transmission is excluded;
fuel prices unchanged. India's costs are lower (scale), so this asks: *is the neighbours' limited
domestic build a geography result, or an artifact of higher (sub-scale) cost assumptions?*

**`B_Opt_IndiaCostsFuel`** — as `IndiaCosts` **plus** non-India fuel-supply `VariableCost`
(`MIN*` techs) set to India's prices. Because India has cheap coal but dearer gas, this is a
*mixed* shift, not a uniform discount.

### 4.6 Validation, coherence & solve settings
- **Machinery neutrality.** `B_Opt_DirBidir` (empty direction map) reproduces `B_Opt_Clipped` to
  the decimal (identical objective, identical output-row count) — proving the direction tooling is
  itself neutral, so any `DirContractual` movement is the *directions*, not the plumbing.
- **Pre-solve coherence.** Every scenario passes `glpsol --check` before solving. Two invariants
  drove the patch design: `Residual ≤ MaxCap` (⇒ backstops disabled by activity block, not
  `MaxCap = 0`) and `MaxCap ≥ Residual + MinCapInv` (⇒ `MinCapInv` clamped to headroom).
- **Solve.** CPLEX 22.1.2, dual simplex, `strip_storage: False`, `delay: True`,
  `cplex_threads = 4`. All 13 solve **optimal with zero backstop generation** and none exceeds the
  BAU cost — i.e. no constraint is economically unreasonable.
- **Pre-CPLEX desk-check.** `desk_check.py` computes a partial energy/cost balance (not dispatch)
  that bounds what the constraints force — e.g. Bangladesh forced-firm generation once imports are
  capped and VRE hits its ceilings — as an independent sanity anchor for the solved numbers.

---

## 5. Results

### 5.1 Headline table

| Scenario | System cost (M USD) | Δ vs Clipped | BGD domestic 2050 | Regional trade 2050 (TWh) | CO₂ 2050 (Mt) | Backstop |
|---|--:|--:|--:|--:|--:|:--:|
| A_Calibrated_BAU | 2,224,446.5 | — | 42.3 % | 331.4 | 3,754.5 | 0 |
| A_Calibrated_BAU_Clipped | 2,224,447.8 | — | 42.3 % | 331.4 | 3,754.5 | 0 |
| B_Optimised_VRE | 2,113,984.5 | — | 39.8 % | 441.5 | 1,769.8 | 0 |
| **B_Opt_Clipped** *(baseline)* | **2,115,081.7** | **0** | **37.6 %** | **438.3** | **1,784.0** | 0 |
| C_Target_VRE | 2,158,340.4 | — | 52.6 % | 320.7 | 2,582.2 | 0 |
| C_Target_VRE_Clipped | 2,149,161.8 | — | 54.9 % | 302.2 | 2,582.4 | 0 |
| **B_Opt_TradeCap15** *(volume)* | 2,126,284.1 | +11,202 (+0.5 %) | **86.7 %** | 242.9 | **1,620.5** | 0 |
| **B_Opt_TxCap150** *(capacity)* | 2,141,786.5 | +26,705 (+1.3 %) | **94.5 %** | 97.1 | 1,666.7 | 0 |
| **B_Opt_DirContractual** *(direction)* | 2,120,330.8 | +5,249 (+0.2 %) | **32.1 %** | 423.5 | 1,773.2 | 0 |
| B_Opt_SolarCapexHi | 2,122,387.8 | +7,306 (+0.3 %) | 37.6 % | 438.3 | 1,784.0 | 0 |
| B_Opt_IndiaCosts | 2,076,619.9 | −38,462 (−1.8 %) | 51.3 % | 392.4 | 1,734.9 | 0 |
| B_Opt_IndiaCostsFuel | 2,076,619.9 | −38,462 (−1.8 %) | 51.3 % | 392.4 | 1,734.9 | 0 |
| B_Opt_DirBidir *(validation)* | 2,115,081.7 | +0 (+0.0 %) | 37.6 % | 438.3 | 1,784.0 | 0 |

*BGD 2050 demand = 485.7 TWh (constant across scenarios).*

### 5.2 The cost of physical VRE realism (clipped − unclipped, per pathway)

| Pathway | Δ system cost | Interpretation |
|---|--:|---|
| A_Calibrated_BAU | +1 M USD (+0.00 %) | BAU barely uses VRE → the clip is a **no-op** (a clean sanity check). |
| B_Optimised_VRE | +1,097 M USD (+0.05 %) | Trims Bangladesh wind (9.16→3 GW) + the Sri Lanka solar overbuild — **negligible**. |
| C_Target_VRE | −9,179 M USD (−0.43 %) | Clipping the unphysical wind *and* re-capping the unmeetable Maldives NDC floor is **cheaper**. |

**Takeaway:** enforcing real physical limits does not meaningfully raise the least-cost frontier,
and on the target pathway it removes phantom obligations that were *inflating* cost. The VRE-
dominated story survives physical realism intact.

### 5.3 The three levers — energy security

Baseline: Bangladesh is **37.6 % domestic / 62.4 % imported** in 2050 (`B_Opt_Clipped`).

| Lever | BGD domestic 2050 | BGD net imports (TWh) | Regional trade (TWh) | Δ cost | Δ CO₂ |
|---|--:|--:|--:|--:|--:|
| Baseline (`B_Opt_Clipped`) | 37.6 % | 303.1 | 438.3 | — | — |
| **Volume** (`TradeCap15`) | **86.7 %** | 64.6 (−79 %) | 242.9 (−45 %) | +0.5 % | **−9.2 %** |
| **Capacity** (`TxCap150`) | **94.5 %** | 26.6 (−91 %) | 97.1 (−78 %) | +1.3 % | −6.6 % |
| **Direction** (`DirContractual`) | **32.1 %** | 329.8 (+9 %) | 423.5 (−3 %) | +0.2 % | −0.6 % |

- **Volume and capacity caps buy realism cheaply.** Both pull Bangladesh from 38 % to 87–95 %
  self-sufficient for **under 1.3 %** of system cost, with **no unserved energy**. Bangladesh
  responds by building domestic firm + solar (coal, oil, nuclear headroom; solar already at its
  10.5 GW economic level) — see [Appendix B](#appendix-b--bangladesh-2050-capacity-by-scenario).
- **Self-sufficiency lowers emissions.** Displacing coal-heavy imported power with domestic supply
  makes `TradeCap15` the **lowest-CO₂ scenario of all thirteen** (1,620 Mt, −9.2 %). Energy
  security and decarbonisation align here rather than trading off.
- **Direction deepens dependence — the counter-intuitive result.** Forcing corridors to their real
  one-way flows makes Bangladesh *more* import-reliant (38 %→32 %). Blocking Bangladesh's
  post-2040 export-back to India East removes its incentive to over-build domestic generation for
  export — exposing that part of B_Opt's apparent "self-supply" was really **build-for-export**,
  not genuine domestic security.

### 5.4 Cost & fuel-price robustness

- **`SolarCapexHi` (+10 % solar CapEx) — the VRE result is robust.** System cost rises only +0.3 %
  and the capacity mix is **essentially unchanged** (solar 878.4 GW, Bangladesh 37.6 % domestic —
  both identical to baseline to the reported precision; VRE ceilings stay idle). Even +10 % dearer,
  solar remains far below firm alternatives, so it is still built to the same levels. The
  least-cost pathway's reliance on cheap solar is **not fragile** to a plausible cost shock.
- **`IndiaCosts` — the largest swing in the set, and it is an assumption, not a technology.**
  Harmonising the smaller nations' capital + fixed costs down to India's cuts system cost **−1.8 %**
  (−38,462 M USD, the biggest move of any lever) and lifts Bangladesh to **51.3 % domestic**
  (from 37.6 %), as cheaper capital lets it self-build firm capacity (Bangladesh nuclear rises to
  11.6 GW, coal to 8.3 GW). **Implication:** a meaningful part of Bangladesh's modelled import
  reliance is a **cost-assumption artifact** (neighbours modelled as more expensive than India),
  not pure geography or resource endowment.
- **`IndiaCostsFuel` ≡ `IndiaCosts` exactly.** Adding fuel-price harmonisation changes nothing at
  the optimum (identical objective and mix) — the non-India fuel-price differentials are **not
  binding** on the solution. A clean negative result worth recording.

### 5.5 System capacity mix, emissions & VRE-ceiling contact
- **Decarbonisation gradient (2050 coal / CO₂):** BAU 506.6 GW / 3,755 Mt → C_Target 378.4 GW /
  2,582 Mt → B_Opt 298.9 GW / 1,784 Mt. The optimised pathway roughly **halves** BAU emissions;
  the interconnection levers shave a further 1–9 % off B_Opt.
- **India dominates absolute build** (solar ~319 GW each in INDNO/INDSO; INDSO wind ~249 GW) and is
  largely insensitive to the Bangladesh-focused levers — the levers redistribute *Bangladesh's*
  supply, not the region's.
- **Which ceilings bind.** After clipping, the physical potential is the active limit for
  `PWRSPVLKAXX` (16 GW), `PWRWONBGDXX` (3 GW), `PWRSPVMDVXX` (1 GW), `PWRSPVBTNXX` (1.81 GW) —
  at 100 % of ceiling across scenarios. India's large nodes retain non-binding headroom (build ≪
  ceiling). This confirms the ceiling is doing real work at exactly the nodes flagged as
  overbuilt, and nowhere spurious.

---

## 6. Implications

1. **Bangladesh's modelled 62 % import dependence is a design choice, not a necessity.** It can be
   reduced to realistic levels (85–95 % self-sufficient) for **≤1.3 %** of system cost with no
   reliability penalty. Energy-security realism is affordable.
2. **The policy instrument matters as much as the target.** A *volume* cap and a *capacity* cap
   both deliver security and cut emissions; a *direction* correction (honouring real contracts)
   does the opposite and **deepens** dependence. Any advice to "constrain trade" must specify
   which margin.
3. **Security and decarbonisation are aligned here.** Domestic self-supply displaces coal-heavy
   imports, so the security levers *lower* CO₂ — they are not a security-vs-climate trade-off.
4. **The result is robust to technology cost, sensitive to relative-cost assumptions.** A +10 %
   solar shock is immaterial; the India-vs-neighbour cost gap is the biggest single driver of both
   system cost and Bangladesh self-sufficiency. Data effort is better spent narrowing the
   neighbours' cost estimates than re-litigating solar CapEx.
5. **Part of B_Opt's "domestic" Bangladesh generation is export-oriented.** The direction lever
   shows it evaporates once export-back is blocked — a caution against reading raw B_Opt domestic
   shares as sovereign energy security.

---

## 7. Relevance: why each sensitivity was chosen

| Sensitivity | Real-world driver it tests |
|---|---|
| **VRE ceiling / clipped baselines** | B_Opt builds VRE beyond atlas-realisable potential; a pathway used for policy must respect physical limits. Prices the cost of that realism. |
| **TradeCap15 (volume)** | B_Opt's 62 % Bangladesh import share vs ~10 % today and ~15–20 % in Bangladesh's own plans — a direct energy-security concern for the region's sole large importer. |
| **TxCap150 (capacity)** | The physical-wires form of the same concern — you cannot import beyond corridor capacity. Answers Zixuan's "disallow / limit the interconnectors" ask. |
| **DirContractual (direction)** | Real corridors are governed by one-way contracts and assets (Adani Godda PPA, Bheramara HVDC, Bhutan hydro export). B_Opt's free bidirectionality overstates Bangladesh's export optionality; this tests faithfulness to contracted reality. |
| **SolarCapexHi** | "VRE is cheap" is the assumption the entire optimised pathway rests on — its fragility must be quantified. |
| **IndiaCosts / IndiaCostsFuel** | The neighbours' capital costs are modelled above India's; tests whether Bangladesh's import reliance is geography or a cost-assumption artifact, and (fuel variant) whether fuel-price gaps bind. |

Every non-inferred interconnector direction is cited to a primary source; the two conceptual
Maldives corridors are flagged as *inferred* and are inactive in B_Opt regardless.

---

## 8. Caveats & limitations

- **Unclipped-B_Opt capital figure is a known artifact.** In the raw outputs, `B_Optimised_VRE`
  Capital (NPV) reads 6,893,228 M USD — ~10.5× every other scenario (all ~630–690k, including
  `B_Opt_Clipped` at 653,246). It appears to be an aggregation glitch for that one scenario cell;
  it does **not** affect any reported System cost or delta (all benchmarked against the clean
  `B_Opt_Clipped`). *Recommend a data check on the concat for that cell before publishing any
  capital/opex split.*
- **`TxCap150` blocks zero-residual corridors.** The 1.5×Residual rule sets a zero cap on corridors
  with no 2023 residual (BTN→BGD, IND-S→LKA) that B_Opt used ~9 and ~8 GW of. A softer
  "TxCap150-Pipeline" variant (floor those at their evidence pipeline) is noted but **not built**.
- **Conceptual corridors (Maldives).** `TRNMDVXXINDSO`, `TRNLKAXXMDVXX` have no committed project;
  directions are inferred from Maldives' deficit-only status and are non-binding (0 build in B_Opt).
- **Node-mapping.** Dhalkebar–Muzaffarpur physically lands in India's Eastern region (Bihar), so it
  arguably belongs to `TRNINDEANPLXX` not `TRNINDNONPLXX`. Both are bidirectional → no effect on
  current results; noted for a possible separate node-mapping review.
- **Desk-check is a partial balance, not dispatch.** It bounds what constraints force; all headline
  numbers come from full CPLEX solves.
- **Superseded configs remain on disk.** Early drafts (`B_Opt_TradeCap50`, `B_Opt_SolarHi10`,
  `B_Opt_LinkFreeze`, `B_Opt_TradeCap30`) are retained for provenance but are **not** part of the
  analysed 13-scenario set.
- **System cost convention.** Reported cost = CPLEX objective + a 157,222 M USD constant term
  excluded by CPLEX's `Objective =` line. Applied uniformly, so deltas are unaffected.

---

## 9. Reproducibility

Full build steps, the file-edit trail, and the per-scenario workflow are in
[`PHASE_B_IMPLEMENTATION_LOG.md`](../../docs/archive/phase-b/PHASE_B_IMPLEMENTATION_LOG.md). In brief (conda env `OSTRAM-env`;
always `set PYTHONIOENCODING=utf-8 && chcp 65001` first):

```bash
# 1. (re)generate patches.json for the computed scenarios
python sensitivity_expansion\gen_sensitivity_patches.py
# 2. build patched A-O = source A-O + shared VRE ceiling layer
python sensitivity_expansion\apply_patches.py --scenario <SCEN> --source-scenario B_Optimised_VRE
# 3. DIRECTION scenarios only — overlay flow-direction edits on the AR files
python A3_process\rules_scripts\set_interconnector_direction.py --input-dir A1_Outputs\A1_Outputs_<SCEN> --yaml A3_process\rules_scripts\configs\<SCEN>\set_interconnector_direction.yaml
# 4. compile → 5. verify (glpsol --check) → 6. solve (CPLEX) → 7. analyse
python B1_Run_Compiler.py --scenarios <SCEN>
python B2_Executing_OG_Model.py --scenarios <SCEN>
python ..\tools\analysis\concat_all_scenarios.py
python ..\tools\analysis\analyse_sensitivity.py  # -> sensitivity_report.txt / sensitivity_comparison.csv
```

Validation of the design (pre-CPLEX) is in `validate_sensitivity_configs.py` (7 checks) and
`desk_check.py`; both write to `reports/`.

---

## Appendix A — Full metric matrix (13 scenarios)

2050 values unless noted. Source: `sensitivity_comparison.csv` (2026-07-09).

| Metric | A_BAU | A_BAU_Clip | B_Opt | **B_Opt_Clip** | C_Target | C_Tgt_Clip | TradeCap15 | SolarCapexHi | TxCap150 | IndiaCosts | IndiaCostsFuel | DirBidir | DirContractual |
|---|--:|--:|--:|--:|--:|--:|--:|--:|--:|--:|--:|--:|--:|
| System cost (M USD) | 2,224,446 | 2,224,448 | 2,113,984 | **2,115,082** | 2,158,340 | 2,149,162 | 2,126,284 | 2,122,388 | 2,141,786 | 2,076,620 | 2,076,620 | 2,115,082 | 2,120,331 |
| Coal (GW) | 506.6 | 506.6 | 298.8 | 298.9 | 378.4 | 373.8 | 298.4 | 298.9 | 312.7 | 299.4 | 299.4 | 298.9 | 303.6 |
| Solar (GW) | 282.8 | 282.8 | 886.6 | 878.4 | 813.1 | 812.7 | 878.4 | 878.4 | 870.1 | 878.4 | 878.4 | 878.4 | 870.1 |
| Wind (GW) | 142.3 | 142.3 | 506.5 | 497.9 | 195.9 | 230.0 | 497.7 | 497.9 | 497.5 | 497.8 | 497.8 | 497.9 | 498.0 |
| Storage (GW) | 94.6 | 94.6 | 68.7 | 68.7 | 83.3 | 75.4 | 68.7 | 68.7 | 68.7 | 68.7 | 68.7 | 68.7 | 68.7 |
| CO₂ (Mt) | 3,754.5 | 3,754.5 | 1,769.8 | 1,784.0 | 2,582.2 | 2,582.4 | 1,620.5 | 1,784.0 | 1,666.7 | 1,734.9 | 1,734.9 | 1,784.0 | 1,773.2 |
| BGD domestic gen (TWh) | 205.4 | 205.4 | 193.5 | 182.6 | 255.3 | 266.6 | 421.1 | 182.6 | 459.1 | 249.2 | 249.2 | 182.6 | 155.9 |
| BGD net imports (TWh) | 280.3 | 280.3 | 292.2 | 303.1 | 230.4 | 219.1 | 64.6 | 303.1 | 26.6 | 236.5 | 236.5 | 303.1 | 329.8 |
| BGD domestic share | 42.3 % | 42.3 % | 39.8 % | 37.6 % | 52.6 % | 54.9 % | 86.7 % | 37.6 % | 94.5 % | 51.3 % | 51.3 % | 37.6 % | 32.1 % |
| BGD solar (GW) | 2.4 | 2.4 | 10.5 | 10.5 | 27.9 | 27.9 | 10.5 | 10.5 | 10.5 | 10.5 | 10.5 | 10.5 | 10.5 |
| Cross-border trade (TWh) | 331.4 | 331.4 | 441.5 | 438.3 | 320.7 | 302.2 | 242.9 | 438.3 | 97.1 | 392.4 | 392.4 | 438.3 | 423.5 |
| Backstop gen (TWh) | 0.0 | 0.0 | 0.0 | 0.0 | 0.0 | 0.0 | 0.0 | 0.0 | 0.0 | 0.0 | 0.0 | 0.0 | 0.0 |

*Capital/OpEx split omitted pending the unclipped-B_Opt capital data check (§8).*

## Appendix B — Bangladesh 2050 capacity by scenario

Installed capacity (GW), node BGDXX. Shows how Bangladesh re-supplies itself under each lever.

| Family | B_Opt_Clip | TradeCap15 | TxCap150 | DirContractual | IndiaCosts |
|---|--:|--:|--:|--:|--:|
| PWRCOA (coal) | 6.61 | 7.10 | 6.94 | 5.84 | **8.32** |
| PWRNGS (gas) | 30.31 | 31.32 | 31.22 | 28.07 | 30.31 |
| PWROIL (oil) | 15.28 | 12.73 | 13.04 | 13.79 | 14.73 |
| PWRSPV (solar) | 10.50 | 10.50 | 10.50 | 10.50 | 10.50 |
| PWRWON (wind) | 3.00 | 3.00 | 3.00 | 3.00 | 3.00 |
| PWRHYD (hydro) | 0.83 | 0.83 | 0.83 | 0.83 | 0.83 |
| PWRURN (nuclear) | 3.41 | 3.88 | 3.70 | 2.40 | **11.59** |
| PWRSDS (short storage) | 4.30 | 4.30 | 4.30 | 4.30 | 4.30 |
| PWRLDS (long storage) | 2.60 | 2.60 | 2.60 | 2.60 | 2.60 |

*Solar and wind are pinned at their ceiling / economic level across levers; the security levers are
met by **firm** capacity (coal, oil, nuclear). `IndiaCosts` shows the biggest domestic build —
cheaper capital unlocks Bangladesh nuclear (3.4→11.6 GW) and coal (6.6→8.3 GW).*

## Appendix C — Interconnector direction map

`B_Opt_DirContractual` — 9 corridors locked one-way, 2 left bidirectional. Full citations in
[`reference/interconnector_direction_references.md`](reference/interconnector_direction_references.md).

| Corridor | Direction kept | Basis (short) | Confidence |
|---|---|---|---|
| TRNBGDXXINDEA | India (ER) → Bangladesh | Adani Godda PPA; Bheramara HVDC | High |
| TRNBGDXXINDNE | India (NER) → Bangladesh | Tripura–Comilla 160 MW since 2016 | High |
| TRNBTNXXINDEA | Bhutan → India (ER) | Tala/Chukha/Mangdechhu hydro export | High |
| TRNBTNXXINDNE | Bhutan → India (NER) | Kurichhu–Salakati wet-export | High |
| TRNBTNXXBGDXX | Bhutan → Bangladesh | Dorjilung trilateral (planned) | Med |
| TRNNPLXXBGDXX | Nepal → Bangladesh | Oct-2024 tripartite deal | High |
| TRNINDSOLKAXX | India → Sri Lanka | Madurai–Mannar HVDC; import-dominant per CEB LTGEP | High |
| TRNMDVXXINDSO | India → Maldives | Conceptual; deficit island → import only | High (inferred) |
| TRNLKAXXMDVXX | Sri Lanka → Maldives | No project exists; deficit island → import only | Inferred |
| TRNINDEANPLXX | Nepal ↔ India (ER) | Genuinely seasonal | High |
| TRNINDNONPLXX | Nepal ↔ India (NR) | Genuinely seasonal | High |

---

*CLG · OSTRAM — generated 2026-07-09. Numbers from the 2026-07-09 combined solve; methodology and
build trail cross-referenced to `docs/archive/phase-b/PHASE_B_IMPLEMENTATION_LOG.md`.*
