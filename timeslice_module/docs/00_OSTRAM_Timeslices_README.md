# OSTRAM Timeslices — Read Me First

**From:** Luis (CLG)
**For:** Natalia
**Date:** 2026-06-15
**Purpose:** Everything you need to fold the OSTRAM timeslice work into the project documentation.

---

## 1. The one-paragraph version

This package contains the full timeslice pipeline for the OSTRAM (OSeMOSYS South Asia, UN ESCAP) model — the scripts that generate it, the methodology behind the chosen scheme, the node-by-node data provenance, and the output workbook the model consumes. **Almost all the methodology is already written** (Section 4 maps it). Your task is to consolidate it into our formal documentation. Before you do, read Section 3 (the adopted scheme) and Section 5 (corrections) — there are a few stale details in the older docs that must not make it into the final write-up.

---

## 2. The bundle — three pieces

You've been given three things. They serve different purposes:

| Piece | What it is | Use it for |
|---|---|---|
| **`OSTRAM_Timeslice_Outputs.xlsx`** | The model-ready output: demand fractions + capacity factors per node, per timeslice. | The actual numbers the model runs. This is the artifact the documentation describes. |
| **The zip (this package)** | Scripts, docs, small reference data, per-node output CSVs. | Reading and understanding. **Not** a cold-rebuild bundle — the bulk raw data was excluded for size (see below). |
| **The folder link** | The complete working directory (`asia_ostram_data`). | Only if you need to *re-run* the pipeline. It has the raw inputs the zip omits: the ~1,140 PGCB daily files, the PUCSL raw files, the Renewables.ninja CSVs. |

**Why the zip can't rebuild on its own:** to keep it small, the raw demand/CF source files were filtered out. The pipeline scripts are all here and the per-node *outputs* are all here, but a from-scratch re-run needs the raw inputs from the folder link. For documentation you don't need to re-run anything — everything you need to describe is in the zip.

---

## 3. The adopted timeslice scheme — state this exactly

The model uses **4 seasons × 5 dayparts = 20 timeslices**.

| Season | Months | Days |
|---|---|---|
| S1 Winter | Dec–Feb | 90 |
| S2 Pre-monsoon | Mar–May | 92 |
| S3 SW Monsoon | Jun–Sep | 122 |
| S4 Post-monsoon | Oct–Nov | 61 |

| Daypart | Hours (local) |
|---|---|
| D1 Night | 00–06 |
| D2 Solar window | 06–17 |
| D3 Evening peak | 17–20 |
| D4 Late evening | 20–22 |
| D5 Late night | 22–24 |

Scheme ID: **`5dp_D_6_17`**.

**Important nuance for the write-up:** the sensitivity analysis (see `ostram_timeslice_handover.md`) identifies `6dp_D_ramp` (24 timeslices) as the *unconstrained* winner and presents it as "the decision." **We adopted the 5-daypart scheme `5dp_D_6_17` instead, deliberately, to keep the model within the solver's variable/constraint budget.** `5dp_D_6_17` is the best-scoring scheme at a 5-daypart budget and is the documented fall-back in that same analysis. So the framing in the documentation should be: *6dp_D_ramp is the methodologically optimal scheme; 5dp_D_6_17 is the adopted scheme under the compute constraint, and the recommended upgrade path if solver capacity increases.* Do **not** describe the model as running 24 timeslices — it runs 20.

---

## 4. Document map — which file gives you what

| File | What it contains | Caveat |
|---|---|---|
| `ostram_timeslice_handover.md` | **The selection methodology.** Sensitivity sweep, v1-vs-v2 ranking, why these boundaries, robustness checks, known weaknesses, and a ready-to-paste methods paragraph (its Section 7). | Section 7 is written for **6dp/24 timeslices**. Use the **5dp-adapted version in Section 6 below** instead. The doc leads with 6dp as "the decision" — re-frame per Section 3 above. |
| `OSTRAM_Timeslice_System_Summary.md` | **The data provenance.** Node-by-node sources (PGCB, PUCSL, Grid-India, CEA, NDOR, BPC, Ninja), the run sequence, the design decisions. | Has stale details — see Section 5. The node-by-node provenance is good; the scheme/MDV/path details are not. |
| `SCHEMES.md` | Design rationale for every candidate scheme in the sweep. | Background/appendix material. Current. |
| `ostram_seasonality_methodnote.md`, `ostram_blockmean_methodnote.md` | Method detail on seasonality definition and block-mean CF construction. | Current. |
| Per-country `Verification_CFs_Dems/` folders + country READMEs | Per-node verification of the demand fractions and capacity factors against source data. | The audit trail for each country's numbers. |
| `_sensitivity/figures/` (103 figures) | Publication-ready plots. Handover doc Section 10 lists the priority set for a deck. | Figures reflect the full sweep, including both 5dp and 6dp schemes. |

---

## 5. Corrections — do NOT inherit these

The older docs predate the final state. Fix these in the documentation:

1. **Scheme.** `OSTRAM_Timeslice_System_Summary.md` shows an old **4-daypart** config block (00-06 / 06-12 / 12-18 / 18-24). That is neither the adopted scheme nor the methodological winner — it is a pre-sweep default. The adopted scheme is **`5dp_D_6_17`** (Section 3).

2. **Maldives.** The system summary lists MDV as *excluded* ("What's NOT in this script"). **MDV is now fully included** — it has demand fractions and CFs in the output workbook (`MDV_Dem`, `MDV_CF` sheets), built from STELCO/yearbook data.

3. **Script and path names** have changed since the system summary was written:
   - sensitivity script: `run_timeslice_sensitivity.py` → **`sensitivity_timeslice_sweep.py`** (and the v2 ranking is `rank_timeslice_schemes_v2.py`)
   - output location: `Downloads\OSTRAM_Timeslices\` → **`asia_ostram_data\OSTRAM_Timeslices\`**
   - downstream A3 script: `A3_update_csvs_from_datapackage.py` → **`A3_update_OSTRAM_WV.py`**

4. **Decision framing.** `ostram_timeslice_handover.md` states the decision as 6dp_D_ramp. Re-frame per Section 3: 6dp is the optimum, 5dp is what we run.

---

## 6. Drop-in methods paragraph (5-daypart version)

The handover doc's Section 7 paragraph is written for 6dp/24 timeslices. Here is the equivalent for the adopted **5dp_D_6_17 / 20-timeslice** scheme. Every figure below traces to `_sensitivity/ranking_v2/ranking_report.txt`.

> The OSTRAM model uses four seasons (Winter, Pre-monsoon, SW Monsoon, Post-monsoon) crossed with five dayparts per day, for a total of 20 timeslices per year. The daypart boundaries (00–06, 06–17, 17–20, 20–22, 22–24 local time; scheme `5dp_D_6_17`) were selected from a sensitivity sweep over 23 candidate schemes spanning equal-width, solar-window, morning-ramp, and physics-informed designs, evaluated against regional hourly demand profiles for the seven primary regions (Bangladesh, Sri Lanka, and five Indian sub-regions) and re-aggregated Renewables.Ninja solar and wind capacity-factor series. Schemes were scored on a composite weighted by three independent information axes: demand fit (45%, split across step-reconstruction RMSE, peak-preservation ratio, and worst-region RMSE), solar-hour differentiation (45%, the coefficient of variation of block-mean solar capacity factors), and wind-hour differentiation (10%); the three-axis weighting avoids triple-counting the demand signal carried by the correlated RMSE/peak/worst-region metric cluster. Aggregation used capacity weighting across the seven primary regions, with a small defensibility penalty on blocks shorter than three hours. `5dp_D_6_17` is the highest-scoring scheme at a five-daypart budget (composite 0.80, capacity-weighted RMSE 0.032, solar CV 1.94); it isolates the midday solar window (06–17) from the dark hours and resolves the evening peak (17–20) and post-peak taper (20–22, 22–24) as distinct dispatch regimes. The five-daypart resolution was adopted in place of the unconstrained optimum (`6dp_D_ramp`, 24 timeslices) to keep the model within the linear solver's variable and constraint budget; `6dp_D_ramp` remains the recommended upgrade should compute capacity increase.

(If you want the 6dp number for comparison in a footnote: `6dp_D_ramp` scores composite 0.86, RMSE 0.024, solar CV 1.81.)

---

## 7. What to do

1. **Read** `ostram_timeslice_handover.md` (selection methodology) and `OSTRAM_Timeslice_System_Summary.md` (data provenance) — these two cover ~90% of the content.
2. **Apply the corrections** in Section 5 as you write — especially the scheme (5dp, 20 timeslices) and the MDV inclusion.
3. **Use the 5dp methods paragraph** in Section 6 as the timeslice-methodology block of the documentation, rather than the 6dp paragraph in the handover doc.
4. **Pull figures** from `_sensitivity/figures/` as needed; the handover doc's Section 10 lists the priority plots.
5. **Cite provenance** per node from the system summary and the `Verification_CFs_Dems/` folders — every demand fraction and CF traces to a named source (PGCB, PUCSL, Grid-India, CEA, NDOR, BPC, STELCO, Ninja).
6. If anything is unclear or you hit a number you can't trace, flag it to me before it goes in — let's not ship anything unsourced.

---

## 8. If you need to regenerate the outputs

You shouldn't need to for documentation, but for completeness: the pipeline lives in the folder link, not this zip. Run sequence (Spyder, F5):

1. `build_reninja_timeslices.py` (only if the daypart definition changes)
2. `build_ostram_timeslices.py` (~11 s) → writes `OSTRAM_Timeslice_Outputs.xlsx` + per-node CSVs to `asia_ostram_data\OSTRAM_Timeslices\`

The daypart scheme is set by the `DAYPART_DEF` block at the top of both scripts (keep them identical). To reproduce the shipped 5dp file, `DAYPART_DEF` must be the `5dp_D_6_17` boundaries (00-06 / 06-17 / 17-20 / 20-22 / 22-24).

---

*Questions → Luis. Don't ship unsourced numbers.*
