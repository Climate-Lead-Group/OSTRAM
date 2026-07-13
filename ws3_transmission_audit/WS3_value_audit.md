# WS-3 — Transmission Value & Sourcing Audit

**Status:** Phase 0–1 complete · **READ-ONLY** · no model value changed · decision gate below
**Branch:** `validated-baseline-3scenario` (identical commit to `main`)
**Date:** 2026-07-09
**Scenarios audited:** A_Calibrated_BAU · B_Optimised_VRE · C_Target_VRE
**Env:** OSTRAM-env (Python 3.10.20, openpyxl 3.1.5, pandas 2.3.3)

> **UPDATE 2026-07-09 (post-gate):** this is the Phase-0/1 *gate* snapshot; its recommendations were subsequently approved and applied. Interconnector costs are now live (v18 → A3 → compiled); the two `[PENDING]` citations below (NP↔BD, BT↔BD) are **backfilled in `SOASIA_v18_REFS.xlsx`**; internal transmission (§7 / D5) is **calibrated + solved** (per-node residuals, RE 2× CapEx, `OperationalLife` 40). Current state → `WS3_calibration_report.md`. Work now in `OSTRAM_ws3_workcopy_D5`.

---

## 0. Bottom line up front (what you must decide at the gate)

1. **The sourced interconnector costs are NOT the ones the model currently solves with.** The live solved model uses a **legacy distance-computed** cost set (e.g. BD↔IN_E = 292.5 $/kW); the `Interconnector_Params` sheet — which *is* sourced (cost DB + IEEE citations in the REFS workbook) — says 380 $/kW. For **11 of 15** parameterised corridors the live value ≠ the sourced value. **Decision: adopt the sourced v18 values as the model's transmission cost basis?** (Recommended — see §3, §5. The legacy distance model badly misprices submarine links.)

2. **Before that decision can be trusted, the wiring must be confirmed** by a controlled re-run (A2→A3→B1 for BAU) — the definitive test of whether `Interconnector_Params.CapitalCost` even reaches the compile. This is the first Phase-2 action (it regenerates artifacts, so it's past the read-only gate). See §3.

3. **Your sources are in far better shape than the "two flags" note predicted.** REFS already carries full IEEE citations for every interconnector parameter; the cost DB is fully sourced. Submarine values *are* cited (not an evidence gap) but **low-confidence**, and the research says they should go **up**, not down. The genuine citation gaps are only **NP↔BD ($450)** and **BT↔BD ($500)** — and the substitute-value research finds both **defensible** (§6).

4. **Internal transmission stays as-is pending your call:** flat placeholder `100 / 4 / 5`, no RE-vs-non-RE split, `OperationalLife` drift (50 for 4 families, 20 for 2 — YAML says 20). Research supports the `$100 non-RE / $200 RE` calibration when you unlock it. See §7.

Nothing in this list has been applied. All of §9 is a decision list awaiting your sign-off.

---

## 1. Scope, method, deliverables

**Two cost domains (kept distinct):**
- **A. Interconnectors** — 15 parameterised `TRN*****×2` corridor techs in `Interconnector_Params` (+ **3 extra** corridors present in the model but *not* parameterised = 18 total). Priority of this audit.
- **B. Internal transmission** — 6 families × 10 nodes = 60 techs (`RNWTRN/RNWNLI/RNWRPO` RE; `PWRTRN/TRNNLI/TRNRPO` non-RE), injected by `A2_AddTx.py` from `Config_country_codes.yaml`.

**Sources reconciled (per parameter):**
| Layer | File | Role |
|---|---|---|
| LIVE (solved) | `t1_confection/Executables/<scen>_0/<scen>_0_Input.csv` | what CPLEX actually used |
| v18 SOURCED | `t1_confection/A3_process/SOASIA_OSeMOSYS_Template_v18.xlsx` → `Interconnector_Params` | template values (pipeline reads this: `A3_process.py:53`) |
| REFS citations | `SOASIA_v18_REFS.xlsx` → `Interconnector_Params` `IEEE Reference` col | full IEEE citations per (tech, param) |
| Cost DB | `SoAsia_OSTRAM_Cost_Database.xlsx` → `Interconnectors`, `Source_Registry` | sourced origin of the cost values |
| Internal | `Config_country_codes.yaml` | YAML-injected family scalars |

**Non-destructive handling:** both provided workbooks were copied out of `Downloads` (originals untouched) into `ws3_transmission_audit/inputs/`. All model reads only.

**Deliverables (this folder):**
- `verify_base_consistency.py` — Phase-0 acceptance test (runs on existing `.sol`, no re-solve)
- `audit_transmission_values.py` — Phase-1 matrix generator
- `outputs/parameter_source_matrix_<ts>.csv` — 120-row parameter-to-source matrix
- `WS3_value_audit.md` — this report

---

## 2. Phase 0 — base-consistency acceptance test

`verify_base_consistency.py` → **10/12 PASS**, and the 2 "fails" are explained and benign:

| Check | Result |
|---|---|
| `SpecifiedAnnualDemand` identical across scenarios (2023; 2023–2027) | **PASS** |
| `ResidualCapacity` identical (2023; 2023–2027) | **PASS** |
| `CapitalCost` identical (2023; 2023–2027) | **PASS** |
| Objective readable, all 3 scenarios | **PASS** — A=2,224,447 · B=2,113,985 · C=2,158,340 |
| No base-year backstop (feasibility) | **PASS** — BCK 2023 = 0 in all 3 |
| 2023 generation identical A vs B / A vs C | **"FAIL"** — 116/117 techs differ |

**The 2023-generation difference is expected, not a defect.** A follow-up diff shows the *only* 2023 inputs that differ across scenarios are the **constraint overlays** — `TotalAnnualMaxCapacityInvestment`, `TotalAnnualMaxCapacity`, `TotalTechnologyAnnualActivityUpperLimit` (e.g. `PWRCSPINDNO` invest cap A=0.0 vs B=0.1). The calibrated base (demand, residual, capital cost, dispatch economics) is **identical**; scenarios diverge only through their intended VRE/lid caps, which bite from year 1.

**Verdict: the cleaned repo is a valid shared baseline** — the three scenarios share one calibrated base and differ only by policy overlay. The three objective values above are the **WS-3 baseline anchors** for the later cost-impact delta (Phase 4).

**Transmission-specific:** every transmission input (interconnector + internal) is **byte-identical across all three scenarios** — so any transmission recalibration moves all three consistently.

---

## 3. Headline finding — the sourced values are not live

The pre-A3 artifact and the post-A3 solved artifact carry the **same** interconnector CapitalCosts (e.g. 292.487), so **A3 is not currently rewriting interconnector CapitalCost from `Interconnector_Params`.** The live values are non-round and track a **distance × per-km** basis (consistent with the `Ref.km.BY` mechanism in `1_merge_timeslices_into_WV.py` and with CERC/POWERGRID per-km benchmarks); the v18 values are round, sourced point estimates.

**Timestamps (2026-07-07):** inputs built 20:28 → solved 20:36 → **v18 saved 20:44 (after the solve)**. So the sourced round values were entered *after* the last run.

**Two candidate explanations, one test:**
- **(H1) Stale artifacts** — v18 was updated post-run; a re-run *would* apply the sourced values.
- **(H2) Not wired** — the compile derives interconnector CapitalCost from a legacy distance model and never applies `Interconnector_Params.CapitalCost`; a re-run would *not* change them.

Evidence leans toward **H2 / a wiring gap** (pre-A3 == post-A3; 3 corridors exist in the model that aren't in the sheet at all; the FixedCost convention differs — live FOM ≈ 0.35 % of CapEx vs the sheet's 1.5 %). **Definitive test (first Phase-2 action):** re-run `A2→A3→B1` for BAU and re-parse `_Input.csv`; if BD↔IN_E stays 292.5 it's a wiring gap to fix, if it becomes 380 it was stale.

**Why this matters beyond bookkeeping:** the legacy distance model **badly misprices submarine corridors** — it prices India↔Sri Lanka at **508 $/kW** (a subsea HVDC link!) where research says ~1,250. The sourced v18 layer is physically sensible (submarine expensive) and cited. This is the core argument for making the sourced values live.

---

## 4. Provenance — where the interconnector values came from

`SoAsia_OSTRAM_Cost_Database.xlsx` → `Interconnector_Params` (v18) → IEEE citations in `SOASIA_v18_REFS.xlsx`. **13 of 15** parameterised corridors match the cost DB CAPEX/FOM to the cent; the cost DB is itself sourced (`Primary_Source` + URL + `Confidence` + `Source_Registry`). FOM = **1.5 % × CapEx** throughout (a convention, not independent data). REFS supplies per-parameter IEEE citations (CEA TBCB matrix, ADB Bheramara, BPC/JICA, World Bank ETTP, USAID/USTDA, PFCCL, POWERGRID, PGCB, BPDB, IEPMP 2023, CEA Manual on Transmission Planning Criteria 2023).

---

## 5. Interconnector parameter-to-source matrix (CapitalCost, $/kW)

Full per-parameter matrix (CapEx, FOM, life, CF, residual) is in `outputs/parameter_source_matrix_<ts>.csv`. CapitalCost summary:

| # | Tech | Corridor | Segment | LIVE | v18 (sourced) | CostDB | Research (2023$) | Citation | Note |
|--:|---|---|---|--:|--:|--:|---|---|---|
| 1 | TRNBGDXXINDEA | BD↔IN_E | overhead XB | 292.5 | **380** | 380 | 250–380; Bheramara $366–398 | CEA+ADB | v18 defensible; live stale |
| 2 | TRNBGDXXINDNE | BD↔IN_NE | overhead XB | 331.4 | **250** | 250 | ” | CEA+ADB | v18 defensible |
| 3 | TRNBTNXXINDEA | BT↔IN_E | overhead XB | 432.5 | **150** | 150 | BT band ~130–180 | CEA+BPC+JICA | v18 OK; live high |
| 4 | TRNBTNXXINDNE | BT↔IN_NE | overhead XB | **180** | 180 | 180 | ~130–180 | CEA+BPC+JICA | live == v18 ✓ |
| 5 | TRNINDNONPLXX | NP↔IN_N | overhead XB | 488.1 | **100** | 100 | NP $100–372 | CEA+WB ETTP | v18 low-side but cited |
| 6 | TRNINDEANPLXX | NP↔IN_E | overhead XB | 452.9 | **130** | 130 | ” | CEA+WB ETTP | v18 defensible |
| 7 | TRNINDSOLKAXX | LK↔IN_S | **submarine** | 507.8 | **1031** | 1031 (Low) | **~1250 (1000–1600)** | USAID/USTDA 2012 | live badly wrong; raise v18 → ~1250 |
| 8 | TRNINDEAINDNO | IN_N↔IN_E | India-internal | 608.6 | **300** | 300 | ±800kV HVDC $295–476 | CEA+POWERGRID | v18 defensible |
| 9 | TRNINDNOINDWE | IN_N↔IN_W | India-internal | 573.1 | **420** | 420 | ” | CEA TBCB | v18 defensible |
| 10 | TRNINDEAINDNE | IN_E↔IN_NE | India-internal | 425.1 | **220** | 220 | ” (Siliguri premium) | CEA TBCB | v18 defensible |
| 11 | TRNINDEAINDSO | IN_E↔IN_S | India-internal | 666.6 | **550** | 550 | ” | CEA+POWERGRID | v18 defensible |
| 12 | TRNINDSOINDWE | IN_W↔IN_S | India-internal | 497.0 | **320** | 320 | ” | CEA TBCB | v18 defensible |
| 13 | TRNMDVXXINDSO | MV↔IN_S | **submarine** | **1600** | 1600 | 1600 (Low) | **~2800 (1800–4000 @400MW)** | PFCCL 2025 | live == v18; raise → ~2800 |
| 14 | TRNNPLXXBGDXX | NP↔BD | via-India | **450** | 450 | — | **~480 (350–900)** | **[PENDING]** | gap; research → 450 defensible |
| 15 | TRNBTNXXBGDXX | BT↔BD | via-India | **500** | 500 | — | **~550 (400–1000)** | **[PENDING]** | gap; research → 500 defensible |
| + | TRNINDEAINDWE | IN_E↔IN_W | India-internal | 691.4 | *absent* | — | — | — | **not in v18** — add |
| + | TRNINDNEINDNO | IN_NE↔IN_N | India-internal | 645.7 | *absent* | — | — | — | **not in v18** — add |
| + | TRNLKAXXMDVXX | LK↔MV | **submarine** | 507.8 | *absent* | — | subsea → ≫508 | — | **not in v18 + mispriced as overhead** |

**Other parameters (all corridors):** live `OperationalLife` = **50** (India-internal & most XB) / **60** (the 3 extra) vs v18 **40** (XB) / **50** (internal) / **30** (MV) — inconsistent on both sides. Live `FixedCost` ≈ **0.35 % of CapEx**; v18 = **1.5 %**. Live `TransmissionLossFactor`/`CapacityFactor` are sourced in REFS (CEA Manual 2023 / IEPMP 2023) and are not the audit's cost focus.

---

## 6. Substitute-value research (6 parallel sweeps, cited)

**Overhead corridors (India-internal, NP, BT, BD): all bands defensible.** ADB South Asia WP-38 (2015) puts overland cross-border at ~$370–1,000/kW; CERC 2010 + POWERGRID give 400 kV D/C ≈ $0.34 M/km and 765 kV D/C ≈ $115/kW·1000 km; ±800 kV HVDC $295/kW (Champa–Kurukshetra) to $476/kW (Raigarh–Pugalur). Verdict: model overhead values are generous-for-AC, fair-for-HVDC — **keep**.

**Submarine — the values should go UP:**
- **India↔Sri Lanka:** ~**$1,250/kW** (range 1,000–1,600). CEB LTGEP 2023 $1,374/kW (500 MW); Apr-2025 MoU ₹9,900 cr/1,000 MW ≈ $1,190/kW; Crete–Attica analog $1,200–1,350/kW. The sourced 1031 is a defensible 2012 vintage but ~15–25 % low for 2023. Confidence Med-High.
- **India↔Maldives:** ~**$2,800/kW** at 400 MW (range 1,800–4,000; ~1,800–2,200 if re-rated to 1,000 MW). No project-specific data exists; proxies: Tyrrhenian $1,840/kW (deep, 1,000 MW), EuroAsia $2,130/kW, POWERGRID Arabian Sea $1,690–1,930/kW (2,500 MW). The sourced 1600 is untraceable and optimistic. Confidence Low. **Cable cost scales with km not kW → a 700 km/400 MW link is expensive per kW.**

**NP↔BD & BT↔BD (the two uncited corridors):** both **defensible**. Anchor = ADB Bheramara HVDC B2B ~$382/kW (1,000 MW, 2010–18). NP↔BD ~$480 (350–900); BT↔BD ~$550 (400–1,000), correctly higher (longer/mountainous). **Important nuance:** today these flows are **wheeling on the existing Indian grid** (Nepal→BD 40 MW via Muzaffarpur; border 6.4 ¢/kWh → delivered ~Tk 8.17–8.50), with **no dedicated new line** — so $450/$500 is a proxy for a *future* dedicated corridor. IEPMP 2023 raises BD import assumption to 9,000 MW but publishes no transmission CAPEX.

---

## 7. Internal transmission (domain B)

All 6 families, all 10 nodes, **identical across scenarios** — a pure placeholder:

| Param | Value | Note |
|---|---|---|
| CapitalCost | **100** | flat; identical for Maldives (~1 GW) and India-West (~240 GW) |
| FixedCost | **4** | flat |
| ResidualCapacity | **5** | flat placeholder (not per-node) |
| CapacityToActivityUnit | 31.536 | OK |
| OperationalLife | **50** (RNWTRN, PWRTRN, TRNRPO, TRNNLI) / **20** (RNWRPO, RNWNLI) | **DRIFT** — YAML says 20 for all |
| TotalAnnualMaxCapacityInvestment | *(absent)* | the **dead YAML key** confirmed — not compiled for transmission |

No RE-vs-non-RE differentiation, no sourcing. Research (per WS-3 brief: LBNL Gorman-Mills-Wiser 2019; NREL ATB 2024 spur $100/kW; India GEC intra-state $51–73/kW) supports **non-RE ~$100 / RE ~$200 (2×, range 1.5–3×)**. Left locked pending your call (§9-D5).

---

## 8. The decision gate — what I need from you

| # | Decision | Recommendation |
|---|---|---|
| **D1** | Confirm the wiring first: run the controlled A2→A3→B1 BAU re-run test (§3). | **Yes** — prerequisite to everything; it's the first Phase-2 action. |
| **D2** | Make the **sourced v18 values** the model's interconnector cost basis (replace legacy distance-computed values)? | **Yes** — they're cited and physically sensible; the live model misprices submarine links. |
| **D3** | **Submarine**: adopt research values — LK 1031→**~1250**, MV 1600→**~2800** (or re-rate MV to 1,000 MW)? | Raise both; MV needs a rating decision. |
| **D4** | **OperationalLife**: standardise all transmission (interconnector + internal) to **40 yr**? | **Yes** — resolves both drifts; matches cost-DB `Life_yr=40`. |
| **D5** | **Internal RE vs non-RE**: unlock `$100 non-RE / $200 RE` with an exposed multiplier for WS-1 (1.5/2/3×)? | Defer to your word — research-backed and ready. |
| **D6** | **3 unparameterised corridors** (IN_E↔IN_W, IN_NE↔IN_N, LK↔MV): add to `Interconnector_Params` with sourced values (LK↔MV is submarine, currently mispriced ~508)? | **Yes** — add + reprice LK↔MV as subsea. |
| **D7** | **NP↔BD / BT↔BD**: keep 450/500 and backfill citations (research-defensible), or model as wheeling? | Keep + cite; flag wheeling-vs-newbuild as a modelling assumption. |

**Ready-to-paste backfill citations** for the two `[PENDING]` corridors (from the research, IEEE-style):
- **NP↔BD (450 $/kW):** Asian Development Bank, "Bangladesh–India Electrical Grid Interconnection Project (Bheramara HVDC B2B, 2×500 MW)," ADB Projects 44192-013/-016 (~$382/kW, 2010–2018 USD, escalated). Corroborating: World Bank, "Nepal–India Electricity Transmission and Trade Project (P115767)"; tripartite Nepal–India–Bangladesh Power Sales Agreement, 3 Oct 2024.
- **BT↔BD (500 $/kW):** ADB SASEC cross-border corridor cost envelope (WP-38, 2015: overland cross-border $370–1,000/kW); Katihar–Parbatipur–Bornagar 765 kV cross-border corridor (planning; POWERGRID-financed); Bhutan Dorjilung (World Bank P501271, 2025) — noting Bhutan→BD is currently unfinanced/aspirational.

---

## 9. Recommended Phase 2–4 plan (pending your gate sign-off — NOT yet executed)

1. **Phase 2 (confirm):** controlled BAU re-run test → settle H1/H2 (§3). Compute internal residuals per node (peak×1.2, RE/non-RE split) → CSV desk-check.
2. **Phase 3 (inject, only approved values):** if H2, wire `Interconnector_Params` into the compile (or patch the post-A2 snapshot per full tech code); set approved submarine/life values + source text; internal RE/non-RE CapEx via one exposed multiplier; `OperationalLife=40` for all transmission; add the 3 missing corridors.
3. **Phase 4 (re-verify):** re-run Phase 0; confirm 3 scenarios still feasible & base-consistent; report the objective delta vs the §2 anchors (the WS-3 cost-impact number).

---

## Appendix — research sources (substitute values)

CERC *Benchmark Capital Cost for 400/765 kV Transmission Lines* (Order L-1/30/2010); World Bank/ESMAP *Understanding the Cost of Transmission Infrastructure* (2026); CEA *Transmission for 500 GW RE by 2030* (2022) & *Manual on Transmission Planning Criteria* (2023); ADB South Asia WP-38 (Chattopadhyay 2015); ADB Bheramara 44192-013/-016; Siemens/NS Energy (Bheramara block 2); POWERGRID Champa–Kurukshetra & Raigarh–Pugalur ±800 kV; NREL iScience (DeSantis 2021), NREL ATB 2024; CEB LTGEP 2023–2042; India–Sri Lanka HVDC MoU (Apr 2025); Crete–Attica/Tyrrhenian/EuroAsia subsea HVDC; POWERGRID Arabian Sea subsea; Härtel et al. (2017) + ACER 2026 UIC. Full figures, URLs and confidence per claim are in the session research logs.
