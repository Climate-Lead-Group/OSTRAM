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
    def test_empty_cells_filled_with_9999(self, working_dir):
        """Truly empty (None/NaN) MaxInv cells in allowed techs -> 9999."""
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
                if pd.isna(old):
                    assert new == ALLOWED_FILL_VALUE, (
                        f"{tech} {yr}: empty cell should -> {ALLOWED_FILL_VALUE}, got {new}"
                    )

    def test_explicit_zero_cells_preserved(self, working_dir):
        """Explicit 0s in allowed-tech MaxInv must be preserved (encode policy)."""
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
                if not pd.isna(old) and old == 0:
                    new = post_max_inv.loc[tech, yr]
                    assert new == 0, (
                        f"{tech} {yr}: explicit 0 should stay 0, got {new}"
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
class TestProjectionModeFlip:
    """After we write values into a row, Projection.Mode must flip
    EMPTY -> 'User defined' so the values get picked up downstream."""

    def test_zeroed_techs_max_cap_mode_flipped(self, working_dir):
        run(working_dir, sheets=["Secondary Techs"])
        df = _read(working_dir / PARAM, "Secondary Techs")
        rows = df[
            (df["Parameter"] == MAX_CAP_PARAM)
            & (df["Tech"].isin(set(df["Tech"]) - _allowed_techs(working_dir)))
        ]
        assert (rows["Projection.Mode"] != "EMPTY").all(), (
            "Some zeroed-tech MaxCapacity rows still have Projection.Mode=EMPTY"
        )

    def test_zeroed_techs_max_inv_mode_flipped(self, working_dir):
        run(working_dir, sheets=["Secondary Techs"])
        df = _read(working_dir / PARAM, "Secondary Techs")
        rows = df[
            (df["Parameter"] == MAX_INV_PARAM)
            & (df["Tech"].isin(set(df["Tech"]) - _allowed_techs(working_dir)))
        ]
        # Some may have been "User defined" pre-edit — we just need EMPTY gone
        assert (rows["Projection.Mode"] != "EMPTY").all()

    def test_allowed_techs_max_inv_mode_flipped(self, working_dir):
        run(working_dir, sheets=["Secondary Techs"])
        df = _read(working_dir / PARAM, "Secondary Techs")
        rows = df[
            (df["Parameter"] == MAX_INV_PARAM)
            & (df["Tech"].isin(_allowed_techs(working_dir)))
        ]
        assert (rows["Projection.Mode"] != "EMPTY").all(), (
            "Some allowed-tech MaxInv rows still have Projection.Mode=EMPTY"
        )

    def test_flipped_rows_now_say_user_defined(self, working_dir):
        log = run(working_dir, sheets=["Secondary Techs"])
        df = _read(working_dir / PARAM, "Secondary Techs")
        for flip in log["sheets"][0]["projection_mode_flips"]:
            row = df[
                (df["Tech"] == flip["tech"]) & (df["Parameter"] == flip["parameter"])
            ]
            assert len(row) == 1
            assert row.iloc[0]["Projection.Mode"] == "User defined"

    def test_non_empty_modes_left_alone(self, working_dir):
        """Pre-existing non-EMPTY modes (e.g. 'User defined' already set)
        must not be reported as flips."""
        pre = _read(working_dir / PARAM, "Secondary Techs")
        log = run(working_dir, sheets=["Secondary Techs"])
        # Every flip in the log should have come from a row that was EMPTY pre-edit
        for flip in log["sheets"][0]["projection_mode_flips"]:
            pre_row = pre[
                (pre["Tech"] == flip["tech"]) & (pre["Parameter"] == flip["parameter"])
            ]
            assert pre_row.iloc[0]["Projection.Mode"] == "EMPTY", (
                f"{flip['tech']}/{flip['parameter']}: flipped non-EMPTY mode "
                f"({pre_row.iloc[0]['Projection.Mode']})"
            )

    def test_other_parameter_rows_mode_untouched(self, working_dir):
        """Projection.Mode on rows for parameters other than MaxCap/MaxInv
        AND MinCapInvestment must be byte-identical."""
        pre = _read(working_dir / PARAM, "Secondary Techs")
        run(working_dir, sheets=["Secondary Techs"])
        post = _read(working_dir / PARAM, "Secondary Techs")
        # MinCapInvestment now also gets its mode flipped (separate test below);
        # exclude it here so this test focuses on truly-untouched parameters.
        excluded = {MAX_CAP_PARAM, MAX_INV_PARAM, "TotalAnnualMinCapacityInvestment"}
        mask = ~pre["Parameter"].isin(excluded)
        pre_modes = pre.loc[mask, ["Tech", "Parameter", "Projection.Mode"]].reset_index(drop=True)
        post_modes = post.loc[mask, ["Tech", "Parameter", "Projection.Mode"]].reset_index(drop=True)
        pd.testing.assert_frame_equal(pre_modes, post_modes)

    def test_allowed_max_capacity_mode_untouched(self, working_dir):
        """We don't touch MaxCapacity rows for allowed techs, so their mode
        must be unchanged."""
        pre = _read(working_dir / PARAM, "Secondary Techs")
        run(working_dir, sheets=["Secondary Techs"])
        post = _read(working_dir / PARAM, "Secondary Techs")
        allowed = _allowed_techs(working_dir)
        for tech in allowed:
            pre_mode = pre[
                (pre["Tech"] == tech) & (pre["Parameter"] == MAX_CAP_PARAM)
            ]["Projection.Mode"].iloc[0]
            post_mode = post[
                (post["Tech"] == tech) & (post["Parameter"] == MAX_CAP_PARAM)
            ]["Projection.Mode"].iloc[0]
            assert pre_mode == post_mode, (
                f"{tech} MaxCapacity mode changed from {pre_mode} to {post_mode}"
            )

    # --- MinCapInvestment activation ---

    def test_min_cap_inv_with_data_mode_flipped(self, working_dir):
        """Every MinCapInvestment row that had at least one non-null cell
        and mode=EMPTY must now read 'User defined'."""
        pre = _read(working_dir / PARAM, "Secondary Techs")
        year_cols = [c for c in pre.columns if isinstance(c, int)]
        mci = pre[pre["Parameter"] == "TotalAnnualMinCapacityInvestment"]
        had_data = mci[mci[year_cols].notna().any(axis=1)]
        was_empty = had_data[had_data["Projection.Mode"] == "EMPTY"]["Tech"].tolist()
        assert len(was_empty) > 0  # sanity: there should be some

        run(working_dir, sheets=["Secondary Techs"])

        post = _read(working_dir / PARAM, "Secondary Techs")
        for tech in was_empty:
            mode = post[
                (post["Tech"] == tech)
                & (post["Parameter"] == "TotalAnnualMinCapacityInvestment")
            ]["Projection.Mode"].iloc[0]
            assert mode == "User defined", (
                f"{tech} MinCapInv mode still {mode!r}, expected 'User defined'"
            )

    def test_min_cap_inv_all_null_rows_mode_unchanged(self, working_dir):
        """Rows with NO data at all (all NaN) must keep mode=EMPTY — we have
        no business activating an empty constraint row."""
        pre = _read(working_dir / PARAM, "Secondary Techs")
        year_cols = [c for c in pre.columns if isinstance(c, int)]
        mci = pre[pre["Parameter"] == "TotalAnnualMinCapacityInvestment"]
        all_null_techs = mci[
            mci[year_cols].isna().all(axis=1)
        ]["Tech"].tolist()

        run(working_dir, sheets=["Secondary Techs"])

        post = _read(working_dir / PARAM, "Secondary Techs")
        for tech in all_null_techs:
            pre_mode = pre[
                (pre["Tech"] == tech)
                & (pre["Parameter"] == "TotalAnnualMinCapacityInvestment")
            ]["Projection.Mode"].iloc[0]
            post_mode = post[
                (post["Tech"] == tech)
                & (post["Parameter"] == "TotalAnnualMinCapacityInvestment")
            ]["Projection.Mode"].iloc[0]
            assert pre_mode == post_mode, (
                f"{tech} MinCapInv mode changed from {pre_mode} to {post_mode} "
                f"despite row being all-NaN"
            )

    def test_min_cap_inv_year_values_unchanged(self, working_dir):
        """We never write into MinCapInvestment year cells — only the mode."""
        pre = _read(working_dir / PARAM, "Secondary Techs")
        year_cols = [c for c in pre.columns if isinstance(c, int)]
        run(working_dir, sheets=["Secondary Techs"])
        post = _read(working_dir / PARAM, "Secondary Techs")
        pre_vals = pre[pre["Parameter"] == "TotalAnnualMinCapacityInvestment"][
            ["Tech"] + year_cols
        ].reset_index(drop=True)
        post_vals = post[post["Parameter"] == "TotalAnnualMinCapacityInvestment"][
            ["Tech"] + year_cols
        ].reset_index(drop=True)
        pd.testing.assert_frame_equal(pre_vals, post_vals, check_dtype=False)


def _allowed_techs(working_dir):
    """Helper: read pristine file and return the set of allowed techs."""
    # Read from backup if it exists, otherwise from current state
    df = pd.read_excel(working_dir / PARAM, sheet_name="Secondary Techs")
    year_cols = [c for c in df.columns if isinstance(c, int)]
    from add_max_capacity_investment_rule import categorize_techs
    allowed, _ = categorize_techs(df, year_cols)
    return allowed



    def test_other_parameters_in_target_sheet_untouched(self, working_dir):
        """Non-target parameter rows must be byte-identical EXCEPT for the
        Projection.Mode column on MinCapInvestment, which we now flip
        EMPTY -> User defined for rows that have data."""
        pre = _read(working_dir / PARAM, "Secondary Techs")
        run(working_dir, sheets=["Secondary Techs"])
        post = _read(working_dir / PARAM, "Secondary Techs")
        for param in pre["Parameter"].unique():
            if param in (MAX_CAP_PARAM, MAX_INV_PARAM):
                continue
            pre_sub = pre[pre["Parameter"] == param].reset_index(drop=True)
            post_sub = post[post["Parameter"] == param].reset_index(drop=True)
            if param == "TotalAnnualMinCapacityInvestment":
                # Values must be unchanged; mode column is allowed to flip
                pre_sub = pre_sub.drop(columns=["Projection.Mode"])
                post_sub = post_sub.drop(columns=["Projection.Mode"])
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

    def test_pwrhydbgdxx_2023_explicit_zero_preserved(self, working_dir):
        """PWRHYDBGDXX 2023 has explicit 0 (no investment in base year) — must stay 0."""
        pre = _read(working_dir / PARAM, "Secondary Techs")
        pre_row = pre[
            (pre["Tech"] == "PWRHYDBGDXX") & (pre["Parameter"] == MAX_INV_PARAM)
        ].iloc[0]
        # sanity: 2023 is an explicit 0, not NaN
        assert not pd.isna(pre_row[2023]) and pre_row[2023] == 0

        run(working_dir, sheets=["Secondary Techs"])

        post = _read(working_dir / PARAM, "Secondary Techs")
        post_row = post[
            (post["Tech"] == "PWRHYDBGDXX") & (post["Parameter"] == MAX_INV_PARAM)
        ].iloc[0]
        assert post_row[2023] == 0

    def test_transmission_interconnect_pre_2030_zeros_preserved(self, working_dir):
        """TRN* interconnects have 0 from 2023-2029, real value from 2030.
        That timing encodes 'project not online until 2030' — must not be opened."""
        run(working_dir, sheets=["Secondary Techs"])
        post = _read(working_dir / PARAM, "Secondary Techs")
        row = post[
            (post["Tech"] == "TRNINDEAINDNE") & (post["Parameter"] == MAX_INV_PARAM)
        ].iloc[0]
        for yr in range(2023, 2030):
            assert row[yr] == 0, f"TRNINDEAINDNE {yr}: pre-2030 zero clobbered to {row[yr]}"
        for yr in range(2030, 2051):
            assert row[yr] == 2.0, f"TRNINDEAINDNE {yr}: existing 2.0 clobbered to {row[yr]}"

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
