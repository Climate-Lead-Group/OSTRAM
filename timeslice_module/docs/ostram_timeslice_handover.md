# OSTRAM Timeslice Selection — Handover

**Project:** OSTRAM (OSeMOSYS for Transmission — South Asia, UN ESCAP submission)
**Scope of this document:** Everything learned during the timeslice sensitivity analysis and ranking exercise, including what was decided, why, and what the deliverables contain.
**Decision:** 6 dayparts × 4 seasons = **24 timeslices**, with daypart boundaries `0-5 / 5-8 / 8-17 / 17-20 / 20-22 / 22-24` (scheme ID: `6dp_D_ramp`).

---

## 1. What we picked

| Budget | Winner | Blocks | Why |
|---|---|---|---|
| **N ≤ 6 (chosen)** | **`6dp_D_ramp`** | 0-5 / 5-8 / 8-17 / 17-20 / 20-22 / 22-24 | Best balance across all axes; wins under every tested weighting |
| N ≤ 8 | `6dp_D_ramp` | same | Under v2 weights, still beats `8dp_equal` despite having fewer blocks |
| N ≤ 5 | `5dp_D_6_17` | 0-6 / 6-17 / 17-20 / 20-22 / 22-24 | Fallback if compute budget shrinks |
| N ≤ 4 | `4dp_shifted` | 0-5 / 5-11 / 11-17 / 17-22 / 22-24 | Fallback below that |

**Why this scheme:** it places block boundaries at the three physical regime changes in South Asian load — pre-dawn (05), morning ramp end (08), evening peak onset (17), post-peak taper (20, 22) — rather than at arbitrary equal-width cuts. The result is low demand-fit error *and* clean separation of solar-available hours from dark hours, something equal-width schemes cannot achieve at any block count.

---

## 2. The two rankings (v1 and v2)

We ran the ranking exercise twice. Both agree on `6dp_D_ramp` at N ≤ 6. They disagree at N ≤ 8, and the reasoning behind v2 is what makes the methods story defensible.

### v1 — original 5-metric composite

Weights: RMSE 30%, PP 25%, SCV 20%, maxRMSE 15%, WCV 10%. Demand-fit cluster (RMSE + PP + maxRMSE) carried 70%.

| Budget | v1 winner |
|---|---|
| N ≤ 4 | `4dp_shifted` |
| N ≤ 5 | `5dp_D_6_17` |
| N ≤ 6 | `6dp_D_ramp` |
| N ≤ 8 | `8dp_equal` |

### v2 — three-axis composite (the one we use)

Weights restructured by **information axis** rather than metric count:
- **Demand axis (45%)**: RMSE 30% + PP 10% + maxRMSE 5%
- **Solar axis (45%)**: SCV 45%
- **Wind axis (10%)**: WCV 10%
- WBV dropped entirely.

| Budget | v2 winner |
|---|---|
| N ≤ 4 | `4dp_shifted` |
| N ≤ 5 | `5dp_D_6_17` |
| N ≤ 6 | **`6dp_D_ramp`** |
| N ≤ 8 | **`6dp_D_ramp`** |

The key v1 → v2 change: at N ≤ 8, `8dp_equal` drops to #3 because its SCV is 1.23 (poor solar separation) while `6dp_D_ramp` sits at 1.81. Under v2's emphasis on solar alignment, adding more timeslices does not compensate for cutting them in the wrong places.

---

## 3. Why v2 — the empirical justification

### Metric correlations

At the scheme level (capacity-weighted aggregates across primary regions), the raw metrics correlate as follows:

| | WBV | RMSE | PP | SCV | WCV |
|---|---|---|---|---|---|
| RMSE | 0.93 | 1.00 | -0.88 | -0.10 | -0.83 |
| PP | -0.92 | -0.88 | 1.00 | 0.26 | 0.90 |
| WCV | -0.90 | -0.83 | 0.90 | -0.03 | 1.00 |

This reveals three structural facts:

1. **WBV and RMSE rank schemes identically** (Spearman = 0.999). WBV is pure redundancy; dropped.
2. **RMSE, PP, and maxRMSE all measure the same underlying axis** — "how well the step function fits the demand curve" — and correlate at r ≥ 0.88. In v1 they effectively voted three times for one signal.
3. **SCV is orthogonal** (r = -0.10 with RMSE). It measures something genuinely different: whether block boundaries align with the solar availability window (roughly 06-17 local time).

### The key insight — SCV is not monotonic with block count

Critically, SCV does not automatically improve with more blocks:

- `8dp_equal` has SCV = 1.23 (poor)
- `6dp_D_ramp` has SCV = 1.81 (good)
- `6dp_solar_physics` has SCV = 1.19 (poor despite being "solar-aware")

The difference is **where** the cuts go, not how many there are. Equal-width cuts at 0/3/6/9/… fragment the solar window into several similarly-valued daytime blocks, destroying the coefficient of variation. A well-placed cut at 08 and 17 creates one high-CF sun block vs. one near-zero night block — maximum contrast.

### Honest framing for the methods paragraph

We debated whether to claim that "solar errors matter more than demand errors because demand can be shifted." We rejected that framing. OSTRAM treats demand as inelastic within a timeslice (ceteris paribus), so both metric families flow into investment decisions symmetrically. **The v2 reweighting is justified purely by metric independence, not by differential error cost.**

---

## 4. Robustness checks

### Weight perturbation (at N ≤ 6)

`6dp_D_ramp` wins under every tested weight scheme:

| Weight regime | #1 at N ≤ 6 |
|---|---|
| baseline v1 | 6dp_D_ramp |
| RMSE-heavy | 6dp_D_ramp |
| PP-heavy | 6dp_D_ramp |
| CF-heavy | 6dp_D_ramp |
| worst-case-heavy | 6dp_D_ramp |
| equal weights | 6dp_D_ramp |
| no-maxrmse | 6dp_D_ramp |
| v2 three-axis | 6dp_D_ramp |

Also survives: defensibility-penalty multiplier from 0× (no penalty) to 4× (aggressive penalty), demand-split variants (Option A/B/C for the 45% demand bucket), and SCV cap sensitivity from 1.5 to no cap.

### Seasonal robustness (N ≤ 6, v2 weights)

| Season | #1 |
|---|---|
| Annual | 6dp_D_ramp |
| S1 Winter | 6dp_D_ramp |
| S2 Pre-monsoon | 6dp_D_ramp |
| S3 SW Monsoon | 6dp_D_ramp |
| S4 Post-monsoon | 6dp_D_ramp |

**Bottom line:** the N ≤ 6 recommendation is weight-invariant and season-invariant across the full perturbation space tested. The only race that's genuinely close is N ≤ 5 (`5dp_D_6_17` vs `5dp_asym` flip under some weightings).

---

## 5. Regional breakdown for the chosen scheme

Primary regions only, capacity-weighted:

| Region | Weight | RMSE | SCV | Notes |
|---|---|---|---|---|
| INDWE | 31.2% | 0.018 | 1.90 | Best-fit primary region |
| INDNO | 26.4% | 0.028 | 1.80 | Highest absolute contribution. 11am peak underresolved by 8-17 block |
| INDSO | 22.8% | 0.026 | 1.81 | Broad evening peak fit OK |
| INDEA | 10.8% | 0.018 | 1.57 | Clean fit |
| BGD | 6.0% | 0.033 | 1.86 | Morning ramp captured cleanly |
| INDNE | 1.4% | 0.039 | 1.52 | Sharp 18-19 peak; small weight |
| LKA | 1.2% | 0.042 | 1.79 | Sharpest 18-19 peak; negligible weight |

Tertiary regions (treat as shape-fitting scores only, not observed hourly variability):
- NPL*: RMSE 0.060, SCV 1.64
- BTN*: RMSE 0.071, SCV 1.83

**The one weakness to flag:** INDNO's 11am demand peak reaches ~1.10 normalised but the 8-17 daytime block averages over it at ~1.06 — a ~4% midday underestimate. This will slightly smooth INDNO midday thermal dispatch results. Flag proactively when discussing INDNO-specific outputs.

---

## 6. Things we diagnosed and rejected along the way

### `solar_physics` scheme family — genuinely bad

`5dp_solar_physics` (blocks: `0-5 / 5-8 / 8-12 / 12-16 / 16-24`) ranks #18 overall despite decent RMSE (0.031). The problem is a single 8-hour block from 16:00 to 24:00 that smears two partial-sun hours (16-17) into six hours of total darkness. Per-region SCV ≈ 1.03 — essentially no solar separation.

The scheme's design philosophy (split the sun day into morning/peak/shoulder chunks) is in direct tension with the SCV metric (which wants one tight sun block vs one tight dark block). We tested seven improved variants with better boundaries and the best achievable was SCV = 1.42 — still well below `6dp_D_ramp`'s 1.82. The scheme cannot be rescued while keeping its philosophy.

Decision: keep `solar_physics` in the sweep for documentation but don't consider it as a finalist. This is fair for capacity expansion modelling. If OSTRAM were a production cost model with intraday ramping concerns, the verdict might differ.

### WCV downweighting — right decision, wrong reason initially

Initial intuition was "weight WCV higher because wind has seasonal variation and hydro handles the seasons." This was empirically wrong — WCV is identical across all four seasons in the sweep because wind profiles are re-aggregated from annual data. WCV is actually measuring within-day wind variability.

The correct reason to downweight WCV is different: **the signal is compressed**. Across all schemes, WCV ranges only 0.05-0.26 in primary regions. South Asia wind has weak diurnal structure, so no amount of timeslicing creates meaningful separation. The WCV heatmap shows all primary regions washed-out orange regardless of scheme — only NPL* and BTN* (both flagged representative profiles) show any green.

Kept at 10% weight for tiebreaker value, not signal value.

### Overlay preference shift

In v1 figures, the default demand overlay showed 6 schemes including both `solar_physics` variants. In v2 figures, trimmed to 4 schemes: the top 3 v2 candidates plus `8dp_equal` as an upper-bound reference. Dropped `solar_physics` (rank #18), `6dp_B_6_18` (runner-up at SCV but tail-penalty-heavy), and `3dp_legacy` (not needed for the v2 comparison story).

---

## 7. Methods paragraph — drop-in for OSTRAM documentation

> The OSTRAM model uses four seasons (Winter, Pre-monsoon, SW Monsoon, Post-monsoon) crossed with six dayparts per day, for a total of 24 timeslices per year. The daypart boundaries (00–05, 05–08, 08–17, 17–20, 20–22, 22–24 local time) were selected from a sensitivity sweep over 23 candidate schemes spanning equal-width, solar-window, morning-ramp, and physics-informed designs, evaluated against regional hourly demand profiles for the seven primary regions (Bangladesh, Sri Lanka, and five Indian sub-regions) and re-aggregated Renewables.Ninja solar and wind capacity factor series. Schemes were scored on a composite weighted by three independent information axes: demand fit (45%, split across step-reconstruction RMSE, peak preservation ratio, and worst-region RMSE), solar-hour differentiation (45%, measured as the coefficient of variation of block-mean solar capacity factors), and wind-hour differentiation (10%). The three-axis weighting avoids triple-counting the demand signal carried by the correlated RMSE/PP/maxRMSE metric cluster. Aggregation used capacity weighting across the seven primary regions, with a small defensibility penalty on blocks shorter than three hours. The chosen scheme resolves the morning demand ramp (05–08), the midday solar window (08–17), and the evening peak (17–20) as distinct regimes, achieving a capacity-weighted RMSE of 0.024 and a solar CV of 1.81 — outperforming even an 8-daypart equal-width scheme at 25% lower timeslice cost, because equal-width cuts fragment the solar window rather than aligning with its boundaries.

---

## 8. Known weaknesses to flag to stakeholders

1. **INDNO midday peak underresolved.** The 8-17 solar block averages over an 11am demand maximum. INDNO midday-cycling thermal dispatch and PV curtailment estimates will be smoothed. Treat INDNO-specific midday results as directional.
2. **Two 2-hour tail blocks (20-22 and 22-24).** Defensible (they isolate the post-peak taper, which matters for battery discharge windows), but a reviewer could ask for a single 20-24 block. Sensitivity sweep shows that costs ~0.006 RMSE.
3. **SCV cap at 2.0 is a judgment call.** Schemes in the B/C family reach 2.17-2.24 SCV pre-cap. If a reviewer believes those values are real rather than geometric artifacts, the cap could be raised; the ranking at N ≤ 6 is stable either way but the runner-up identity changes.
4. **NPL* and BTN* are representative daily profiles, not observed hourly data.** Do not cite their fit quality as evidence for the scheme. Tiebreakers only.
5. **Ranking uses input-side metrics only.** The assumption that better input fidelity translates to better output fidelity is reasonable for long-run capacity expansion but not directly tested.
6. **Does not resolve: within-hour ramping, weekday/weekend differences, heat-wave peak days.** Standard OSeMOSYS limitations, unrelated to timeslice choice.

---

## 9. Deliverables produced in this workstream

All files in `_sensitivity/`:

**Rankings and analysis:**
- `rank_timeslice_schemes.py` — v1 ranking script (kept for reproducibility)
- `rank_timeslice_schemes_v2.py` — v2 three-axis ranking script (**use this one**)
- `ranking/` — v1 ranking outputs (`ranking_full.csv`, `ranking_by_budget.csv`, `ranking_report.txt`)
- `ranking_v2/` — v2 ranking outputs (same three files)

**Figures:**
- `sensitivity_figures.py` (original)
- `sensitivity_figures_v2.py` — regenerated figure script (**use this one**). Three changes from the original:
  1. `OVERLAY_PREFERENCE` trimmed to top 3 v2 candidates + `8dp_equal` reference (4 lines per overlay plot)
  2. Heatmaps extended from 2 metrics (RMSE, PP) to 4 (RMSE, PP, SCV, WCV), all using the full candidate set
  3. New `generate_solar_overlays()` produces 45 solar-CF plots showing how each candidate scheme reconstructs the true hourly solar CF curve — companion to the demand overlays
  4. ASCII-safe `[*]` marker and ASCII dashes used to avoid Windows matplotlib encoding issues
- `figures/` — 103 figures total:
  - 45 demand overlays (region × season)
  - 45 solar-CF overlays (region × season) — new
  - 5 convergence plots (one per metric)
  - 4 heatmaps (one per principal metric)
  - 1 scheme definitions
  - 1 improvement bars vs legacy
  - 1 regional bars
  - 1 solve scaling
- `figures/table_*.csv` — 3 reference tables (schemes, metrics, acronyms)

**Source data (not modified in this workstream):**
- `sensitivity_timeslice_summary.csv` — 1,035 rows, 23 schemes × 9 regions × 5 seasons
- `sensitivity_ninja_cfs.csv` — per-scheme solar/wind CFs
- `sensitivity_config.json` — block definitions

---

## 10. Recommended plot selection for a deck

In priority order, for a 12-16 slide OSTRAM presentation:

1. **`fig_solar_INDNO_Annual.png`** — the "money plot." Shows the true hourly solar CF curve against 4 candidate schemes' block-averages. You can *see* why 8dp_equal's equal-width cuts fragment the solar window while 6dp_D_ramp's single wide block aligns with the sun envelope. Per-scheme SCV values annotated.
2. **`fig_heatmap_scv.png`** — the SCV heatmap showing all 20 candidates × 9 regions. Immediately visible: green columns (v2 preferred) vs orange columns (v1 over-rewarded). Capacity-weighted row is primary-only, matching the ranking.
3. **`fig_convergence_scv.png`** — shows SCV is NOT monotonic with N. Crucial for defending 6dp_D_ramp > 8dp_equal under v2.
4. **`fig_demand_INDWE_Annual.png`** or **`fig_demand_LKA_Annual.png`** — a demand overlay for a high-capacity region showing 6dp_D_ramp hugging the profile.
5. **`fig_scheme_definitions.png`** — visual reference of all block boundaries. Recommended scheme starred.
6. **`fig_solve_scaling.png`** — theoretical LP solve-time scaling. Surprising asset: 6dp_D_ramp costs 1.8× the current 4dp baseline while 8dp_equal would cost 2.8× for worse results under v2.
7. Appendix: full 103-figure set, ranking CSVs, methods paragraph from section 7.

---

## 11. Open follow-ups (if timeslice decision gets reopened)

- If a reviewer challenges the v2 reweighting: the v1 ranking produces `6dp_D_ramp` at N ≤ 6 too. The N ≤ 8 result changes between v1 (`8dp_equal`) and v2 (`6dp_D_ramp`). Show both and let the metric-independence argument carry the defense.
- If compute budget permits: `8dp_equal` (32 timeslices) is a defensible upgrade giving a further ~23% RMSE reduction. Requires benchmarking the actual solve time, not the theoretical estimate.
- If compute budget shrinks below 24 timeslices: drop to `5dp_D_6_17` (20 timeslices). Do NOT drop below N=5 — the quality cliff from N=5 to N=4 is ~0.12 on the composite, 4× the N=6 → N=5 drop.
- If INDNO midday investment decisions turn out sensitive to the 8-17 block averaging: consider a 6dp variant with an additional midday split (e.g. 0-5 / 5-8 / 8-12 / 12-17 / 17-20 / 20-24). Not tested in the sweep; would need re-ranking.
- `solar_physics` family is retired. If a research question emerges around intraday solar ramping, revisit whether a different metric (not SCV) should drive the ranking for that use case.

---

*Prepared at the end of the timeslice sensitivity workstream. Handover owner: Luis Victor-Gallardo, CLG.*
