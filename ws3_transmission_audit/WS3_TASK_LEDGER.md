# WS-3 Task Ledger — memory refresher

**Updated:** 2026-07-09 · **Work in:** `OSTRAM_ws3_workcopy_D5` (LIVE, D5 phase) · **`OSTRAM_ws3_workcopy` = FROZEN checkpoint @ interconnectors-done (do not edit)** · **`OSTRAM_clean` = read-only original baseline**

> **Folder layout (D5, 2026-07-09):** `OSTRAM_ws3_workcopy_D5` was forked (full byte-for-byte copy) from `OSTRAM_ws3_workcopy` at the interconnector-complete milestone, so the interconnector work stays frozen and rollback-able while D5 (internal transmission) proceeds here. All D5 edits + pipeline re-runs happen in `_D5`.

## One-line status
**Interconnector calibration COMPLETE + verified.** **D5 internal transmission COMPLETE** — all 3 scenarios re-run (A3→B1) + solved (B2): per-node residuals, RE 2× CapEx, life=40; A==B==C inputs; interconnectors intact; base-consistent (10/12); all feasible. Cost: D5 Δ A −0.283% / B +0.037% / C +0.081%; total WS-3 A −0.073% / B +0.220% / C +0.303%. **Remaining: D7 citations (doc-only) + promotion to `OSTRAM_clean` branch.**

---

## ✅ DONE

**Phase 0 — repo sanity / acceptance test** (`verify_base_consistency.py`)
- 3 scenarios share one calibrated base (SpecifiedAnnualDemand / ResidualCapacity / CapitalCost identical @2023–27); all feasible (no base-year backstop).
- Baseline objective anchors: **A 2,224,447 · B 2,113,985 · C 2,158,340**.
- Base-year *generation* differs across scenarios only via investment/activity caps (intended, not a defect).

**Phase 1 — value & sourcing audit (the gate)** (`audit_transmission_values.py`, `parameter_source_matrix_*.csv`, `WS3_value_audit.md`)
- Provenance: `SoAsia_OSTRAM_Cost_Database` → v18 `Interconnector_Params` → `SOASIA_v18_REFS` IEEE citations.
- Research (6 cited streams): overhead corridors defensible; submarine LK ~1250 / MV ~2800; NP↔BD ~480 / BT↔BD ~550 (currently wheeling on the Indian grid); India CEA/POWERGRID per-km; internal-tx per-kW RE premium 1.5–2×; per-node premiums.

**Phase 2 — live-vs-sourced verify**
- **Found the wiring gap:** model used legacy distance-computed interconnector CapEx from `OG_csvs`; v18 `Interconnector_Params` was never consumed. Confirmed by A3 re-run (pre = post).

**Phase 3 — injection / wiring fix**
- `apply_interconnector_costs.py` = new **core A3 stage** (`stage_ws3_interconnector_costs`). v18 `Interconnector_Params` is now the source of truth (CapEx/FOM → Secondary Techs; OperationalLife → Fixed Horizon Parameters). Per-scenario overrides via the `scenario` column. Self-tested.
- Final v18 values: 18 corridors sourced/research; submarine raised (LK 1031→1250, MV 1600→2800); LK↔MV repriced subsea (508→1250); 3 corridors added; **OperationalLife = 40**; FOM = 1.5%×CapEx.
- A3 + B1 re-run for all 3; sourced values confirmed in the compiled otoole inputs.

**Phase 4 — re-verify** (`WS3_calibration_report.md`)
- All 3 re-solved (B2). Cost impact: **A +0.21% · B +0.18% · C +0.22%**. All feasible (0 backstop). Base-year inputs identical across scenarios.

---

## ⏳ PENDING (your decisions — nothing blocking)

**D5 — internal transmission** *(COMPLETE via ONE late A3 stage `stage_ws3_internal_transmission` → `rules_scripts/apply_internal_transmission.py`; self-tested + dry-run + compiled-verified; all 3 solved optimal + feasible + base-consistent)*
- [x] **Per-node residuals injected** — compiled `RNWTRN<node>`/`PWRTRN<node>` ResidualCapacity now per-node (peak×1.2 split, from frozen `rules_scripts/internal_tx_residuals.csv`); RNWNLI/RNWRPO/TRNNLI/TRNRPO = 0. Replaced flat 5. *(Done in the late stage on the delivered `Demand Techs` sheet, not the snapshot — the snapshot stays pristine flat-5 and is the rollback point.)*
- [x] **RE 2× CapEx applied** — compiled RNWTRN/RNWNLI/RNWRPO CapitalCost 200 / FixedCost 8; PWRTRN/TRNNLI/TRNRPO 100 / 4. Live WS-1 slider = `internal_transmission.re_capex_multiplier` (2.0) in `Config_country_codes.yaml`.
- [x] **OperationalLife = 40** for all 6 internal families in compiled `OperationalLife.csv` (was 50/20; the 50 came from the Stage-1 `3_update_ao_from_extensions.py` merge, which is why a late stage — not the snapshot/YAML — is the reliable lever).
- [x] ~~Per-node cost multipliers~~ — **DROPPED** (uniform cost across nodes; study is interties, intra is just accounting).

**D7 — citations** *(documentation, no model effect)*
- [x] **Backfilled** NP↔BD (`TRNNPLXXBGDXX`) + BT↔BD (`TRNBTNXXBGDXX`) IEEE citations into `SOASIA_v18_REFS.xlsx` → `Interconnector_Params` (CapEx + FOM rows; ADB Bheramara / WB ETTP / SASEC WP-38 / Dorjilung); 0 `[Pending]` left; REFS backed up (`*_PRE_D7_*`). 2026-07-09.
- [x] *(polish)* `Source` column added to the live v18 `Interconnector_Params` — 18 corridor tags; CapEx + 20 sheets verified intact; backup kept. This documents all 18 in the live artifact, so the separate "REFS rows for the 3 added corridors" was **superseded/skipped** (redundant).

**Promotion**
- [ ] Bring the validated workcopy into `OSTRAM_clean` on a **dedicated branch** (branch-only, no merge). Branch base = main (== current tip). `git add` specific files only.

---

## Key facts to remember
- **Work in the copy** (`OSTRAM_ws3_workcopy`); the original is the untouched baseline.
- **B2 CPLEX solves = you** (Anaconda Prompt, cplex_threads=4). Claude preps A3/B1 + verifies.
- Env python: `C:\Users\luisfernando\anaconda3\envs\OSTRAM-env\python.exe` (conda not on PATH; prepend env dirs + set UTF-8).
- Deliverables all in `ws3_transmission_audit/`: `WS3_value_audit.md` (gate), `WS3_calibration_report.md`, `parameter_source_matrix_*.csv`, the scripts, and this ledger.
