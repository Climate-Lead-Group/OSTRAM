"""
test_add_max_capacity_investment_rule.py
=========================================

Run with:  pytest test_add_max_capacity_investment_rule.py -v

Each test gets a fresh copy of A1_Outputs_BAU under tmp_path so they don't
interfere with each other.
"""

from __future__ import annotations

import hashlib
import shutil
import sys
from pathlib import Path

import pandas as pd
import pytest

THIS_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(THIS_DIR))

from add_max_capacity_investment_rule import (  # noqa: E402
    ALLOWED_FILL_VALUE,
    MAX_CAP_PARAM,
    MAX_INV_PARAM,
    categorize_techs,
    run,
)

SOURCE_DIR = THIS_DIR / "A1_Outputs" / "A1_Outputs_BAU"
PARAM = "A-O_Parametrization.xlsx"
OTHER_FILES = (
    "A-O_AR_Model_Base_Year.xlsx",
    "A-O_AR_Projections.xlsx",
    "A-O_Demand.xlsx",
)


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------
@pytest.fixture
def working_dir(tmp_path):
    """Fresh copy of A1_Outputs_BAU per test."""
    target = tmp_path / "A1_Outputs_BAU"
    shutil.copytree(SOURCE_DIR, target)
    return target


def _md5(path: Path) -> str:
    return hashlib.md5(path.read_bytes()).hexdigest()


def _read(path: Path, sheet: str) -> pd.DataFrame:
    return pd.read_excel(path, sheet_name=sheet)


# ---------------------------------------------------------------------------
# Backup integrity
# ---------------------------------------------------------------------------
class TestBackup:
    def test_backup_folder_created(self, working_dir):
        log = run(working_dir, sheets=["Secondary Techs"])
        assert Path(log["backup_dir"]).is_dir()

    def test_backup_contains_all_xlsx_files(self, working_dir):
        log = run(working_dir, sheets=["Secondary Techs"])
        backup = Path(log["backup_dir"])
        src_files = sorted(p.name for p in SOURCE_DIR.glob("*.xlsx"))
        bak_files = sorted(p.name for p in backup.glob("*.xlsx"))
        assert src_files == bak_files

    def test_backup_param_file_byte_identical_to_pre_edit(self, working_dir):
        pre_hash = _md5(working_dir / PARAM)
        log = run(working_dir, sheets=["Secondary Techs"])
        bak_hash = _md5(Path(log["backup_dir"]) / PARAM)
        assert bak_hash == pre_hash

    def test_change_log_written(self, working_dir):
        log = run(working_dir, sheets=["Secondary Techs"])
        assert Path(log["log_path"]).is_file()


# ---------------------------------------------------------------------------
# Zeroed techs: both upper bounds locked at 0 across all years
# ---------------------------------------------------------------------------
class TestZeroedTechs:
    def test_max_capacity_zero_for_all_zeroed_techs_all_years(self, working_dir):
        log = run(working_dir, sheets=["Secondary Techs"])
        df = _read(working_dir / PARAM, "Secondary Techs")
        year_cols = [c for c in df.columns if isinstance(c, int)]
        zeroed = log["sheets"][0]["zeroed_techs"]
        for tech in zeroed:
            rows = df[(df["Tech"] == tech) & (df["Parameter"] == MAX_CAP_PARAM)]
            assert len(rows) == 1
            vals = rows.iloc[0][year_cols].fillna(-999)
            assert (vals == 0).all(), (
                f"{tech}: TotalAnnualMaxCapacity has non-zero entries: "
                f"{vals[vals != 0].to_dict()}"
            )

    def test_max_investment_zero_for_all_zeroed_techs_all_years(self, working_dir):
        log = run(working_dir, sheets=["Secondary Techs"])
        df = _read(working_dir / PARAM, "Secondary Techs")
        year_cols = [c for c in df.columns if isinstance(c, int)]
        zeroed = log["sheets"][0]["zeroed_techs"]
        for tech in zeroed:
            rows = df[(df["Tech"] == tech) & (df["Parameter"] == MAX_INV_PARAM)]
            assert len(rows) == 1
            vals = rows.iloc[0][year_cols].fillna(-999)
            assert (vals == 0).all(), (
                f"{tech}: TotalAnnualMaxCapacityInvestment non-zero: "
                f"{vals[vals != 0].to_dict()}"
            )


# ---------------------------------------------------------------------------
# Allowed techs: fill empty/zero with 9999, preserve existing positive values
# ---------------------------------------------------------------------------
class TestAllowedTechs:
    def test_empty_or_zero_cells_filled_with_9999(self, working_dir):
        pre = _read(working_dir / PARAM, "Secondary Techs")
        year_cols = [c for c in pre.columns if isinstance(c, int)]
        pre_max_inv = pre[pre["Parameter"] == MAX_INV_PARAM].set_index("Tech")[year_cols]

        log = run(working_dir, sheets=["Secondary Techs"])
        allowed = log["sheets"][0]["allowed_techs"]

        post = _read(working_dir / PARAM, "Secondary Techs")
        post_max_inv = post[post["Parameter"] == MAX_INV_PARAM].set_index("Tech")[year_cols]

        for tech in allowed:
            for yr in year_cols:
                old = pre_max_inv.loc[tech, yr]
                new = post_max_inv.loc[tech, yr]
                if pd.isna(old) or old == 0:
                    assert new == ALLOWED_FILL_VALUE, (
                        f"{tech} {yr}: empty/0 should -> {ALLOWED_FILL_VALUE}, got {new}"
                    )

    def test_existing_positive_values_preserved(self, working_dir):
        pre = _read(working_dir / PARAM, "Secondary Techs")
        year_cols = [c for c in pre.columns if isinstance(c, int)]
        pre_max_inv = pre[pre["Parameter"] == MAX_INV_PARAM].set_index("Tech")[year_cols]

        log = run(working_dir, sheets=["Secondary Techs"])
        allowed = log["sheets"][0]["allowed_techs"]

        post = _read(working_dir / PARAM, "Secondary Techs")
        post_max_inv = post[post["Parameter"] == MAX_INV_PARAM].set_index("Tech")[year_cols]

        for tech in allowed:
            for yr in year_cols:
                old = pre_max_inv.loc[tech, yr]
                if not pd.isna(old) and old > 0:
                    new = post_max_inv.loc[tech, yr]
                    assert new == old, (
                        f"{tech} {yr}: existing positive {old} clobbered to {new}"
                    )

    def test_max_capacity_unchanged_for_allowed_techs(self, working_dir):
        pre = _read(working_dir / PARAM, "Secondary Techs")
        year_cols = [c for c in pre.columns if isinstance(c, int)]
        pre_max_cap = pre[pre["Parameter"] == MAX_CAP_PARAM].set_index("Tech")[year_cols]

        log = run(working_dir, sheets=["Secondary Techs"])
        allowed = log["sheets"][0]["allowed_techs"]

        post = _read(working_dir / PARAM, "Secondary Techs")
        post_max_cap = post[post["Parameter"] == MAX_CAP_PARAM].set_index("Tech")[year_cols]

        for tech in allowed:
            pre_vals = pre_max_cap.loc[tech].fillna("NA").tolist()
            post_vals = post_max_cap.loc[tech].fillna("NA").tolist()
            assert pre_vals == post_vals, (
                f"{tech}: TotalAnnualMaxCapacity should be unchanged for allowed techs"
            )


# ---------------------------------------------------------------------------
# Non-interference: nothing else gets touched
# ---------------------------------------------------------------------------
class TestNonInterference:
    def test_other_parameters_in_target_sheet_untouched(self, working_dir):
        pre = _read(working_dir / PARAM, "Secondary Techs")
        run(working_dir, sheets=["Secondary Techs"])
        post = _read(working_dir / PARAM, "Secondary Techs")
        for param in pre["Parameter"].unique():
            if param in (MAX_CAP_PARAM, MAX_INV_PARAM):
                continue
            pre_sub = pre[pre["Parameter"] == param].reset_index(drop=True)
            post_sub = post[post["Parameter"] == param].reset_index(drop=True)
            pd.testing.assert_frame_equal(
                pre_sub, post_sub, check_dtype=False, obj=f"Parameter={param}"
            )

    def test_other_sheets_in_param_workbook_untouched(self, working_dir):
        sheets_to_check = [
            "Fixed Horizon Parameters",
            "Primary Techs",
            "Capacities",
            "Yearsplit",
            "DaySplit",
            "VariableCost",
            "Demand Techs",
            "System Parameters",
        ]
        pre = {s: _read(working_dir / PARAM, s) for s in sheets_to_check}
        run(working_dir, sheets=["Secondary Techs"])
        post = {s: _read(working_dir / PARAM, s) for s in sheets_to_check}
        for s in sheets_to_check:
            pd.testing.assert_frame_equal(
                pre[s], post[s], check_dtype=False, obj=f"Sheet={s}"
            )

    def test_other_AO_files_byte_identical(self, working_dir):
        pre = {f: _md5(working_dir / f) for f in OTHER_FILES}
        run(working_dir, sheets=["Secondary Techs"])
        post = {f: _md5(working_dir / f) for f in OTHER_FILES}
        assert pre == post, "Non-Parametrization AO files were modified"

    def test_row_counts_preserved_in_all_sheets(self, working_dir):
        pre_counts = {
            s: len(df)
            for s, df in pd.read_excel(
                working_dir / PARAM, sheet_name=None
            ).items()
        }
        run(working_dir, sheets=["Secondary Techs"])
        post_counts = {
            s: len(df)
            for s, df in pd.read_excel(
                working_dir / PARAM, sheet_name=None
            ).items()
        }
        assert pre_counts == post_counts


# ---------------------------------------------------------------------------
# Categorisation correctness
# ---------------------------------------------------------------------------
class TestCategorization:
    def test_partition_is_complete_and_disjoint(self, working_dir):
        df = _read(working_dir / PARAM, "Secondary Techs")
        year_cols = [c for c in df.columns if isinstance(c, int)]
        allowed, zeroed = categorize_techs(df, year_cols)
        all_techs = set(df["Tech"].dropna().unique())
        assert allowed.isdisjoint(zeroed)
        assert allowed | zeroed == all_techs

    def test_known_zeroed_tech(self, working_dir):
        # PWRCSPINDEA: no residual, no min cap investment -> zeroed
        log = run(working_dir, sheets=["Secondary Techs"])
        assert "PWRCSPINDEA" in log["sheets"][0]["zeroed_techs"]

    def test_known_allowed_tech(self, working_dir):
        # PWRHYDINDNO: 19.756 GW residual hydro -> allowed
        log = run(working_dir, sheets=["Secondary Techs"])
        assert "PWRHYDINDNO" in log["sheets"][0]["allowed_techs"]


# ---------------------------------------------------------------------------
# Specific cell-level expectations
# ---------------------------------------------------------------------------
class TestCellLevelExpectations:
    def test_pwrhydbgdxx_2024_value_preserved_at_0_166(self, working_dir):
        """PWRHYDBGDXX 2024 has 0.166 (real hydro ramp cap) — must not be clobbered."""
        pre = _read(working_dir / PARAM, "Secondary Techs")
        pre_row = pre[
            (pre["Tech"] == "PWRHYDBGDXX") & (pre["Parameter"] == MAX_INV_PARAM)
        ].iloc[0]
        assert abs(pre_row[2024] - 0.166) < 1e-6  # sanity check on test data

        run(working_dir, sheets=["Secondary Techs"])

        post = _read(working_dir / PARAM, "Secondary Techs")
        post_row = post[
            (post["Tech"] == "PWRHYDBGDXX") & (post["Parameter"] == MAX_INV_PARAM)
        ].iloc[0]
        assert abs(post_row[2024] - 0.166) < 1e-6

    def test_pwrhydbgdxx_2023_zero_filled_to_9999(self, working_dir):
        """PWRHYDBGDXX 2023 was 0 — should be filled with 9999."""
        run(working_dir, sheets=["Secondary Techs"])
        post = _read(working_dir / PARAM, "Secondary Techs")
        post_row = post[
            (post["Tech"] == "PWRHYDBGDXX") & (post["Parameter"] == MAX_INV_PARAM)
        ].iloc[0]
        assert post_row[2023] == ALLOWED_FILL_VALUE

    def test_pwrcspindea_max_cap_zero_in_every_year(self, working_dir):
        """PWRCSPINDEA had 79.1 in every year for MaxCapacity — must be zeroed."""
        pre = _read(working_dir / PARAM, "Secondary Techs")
        pre_row = pre[
            (pre["Tech"] == "PWRCSPINDEA") & (pre["Parameter"] == MAX_CAP_PARAM)
        ].iloc[0]
        year_cols = [c for c in pre.columns if isinstance(c, int)]
        assert (pre_row[year_cols].fillna(0) > 0).any()  # sanity

        run(working_dir, sheets=["Secondary Techs"])

        post = _read(working_dir / PARAM, "Secondary Techs")
        post_row = post[
            (post["Tech"] == "PWRCSPINDEA") & (post["Parameter"] == MAX_CAP_PARAM)
        ].iloc[0]
        assert (post_row[year_cols].fillna(-1) == 0).all()


# ---------------------------------------------------------------------------
# Idempotency and change-log self-consistency
# ---------------------------------------------------------------------------
class TestIdempotency:
    def test_second_run_reports_no_changes(self, working_dir):
        run(working_dir, sheets=["Secondary Techs"])
        log2 = run(working_dir, sheets=["Secondary Techs"])
        sheet_log = log2["sheets"][0]
        assert sheet_log["changes_allowed_techs"] == []
        assert sheet_log["changes_zeroed_techs"] == []

    def test_second_run_produces_identical_file(self, working_dir):
        run(working_dir, sheets=["Secondary Techs"])
        df1 = _read(working_dir / PARAM, "Secondary Techs")
        run(working_dir, sheets=["Secondary Techs"])
        df2 = _read(working_dir / PARAM, "Secondary Techs")
        pd.testing.assert_frame_equal(df1, df2, check_dtype=False)


class TestChangeLogConsistency:
    def test_every_logged_change_matches_post_state(self, working_dir):
        log = run(working_dir, sheets=["Secondary Techs"])
        df = _read(working_dir / PARAM, "Secondary Techs")
        sheet_log = log["sheets"][0]

        for change in sheet_log["changes_allowed_techs"]:
            rows = df[
                (df["Tech"] == change["tech"])
                & (df["Parameter"] == change["parameter"])
            ]
            assert len(rows) == 1
            assert rows.iloc[0][change["year"]] == change["new"]

        for change in sheet_log["changes_zeroed_techs"]:
            rows = df[
                (df["Tech"] == change["tech"])
                & (df["Parameter"] == change["parameter"])
            ]
            assert len(rows) == 1
            assert rows.iloc[0][change["year"]] == 0

    def test_every_preserved_value_still_there(self, working_dir):
        log = run(working_dir, sheets=["Secondary Techs"])
        df = _read(working_dir / PARAM, "Secondary Techs")
        for p in log["sheets"][0]["preserved_existing_values"]:
            rows = df[(df["Tech"] == p["tech"]) & (df["Parameter"] == p["parameter"])]
            assert abs(rows.iloc[0][p["year"]] - p["value"]) < 1e-9
