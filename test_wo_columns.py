"""Tests for WO spacing, New_WO values, week_day (0-6), week_number (per employee)."""

from __future__ import annotations

import os

import pandas as pd
import pytest

from process_wo_columns import add_new_timing_columns, build_wo_columns, validate_wo_output


def _base_df(rows: list[dict]) -> pd.DataFrame:
    for r in rows:
        r.setdefault("Shift", "AS")
        r.setdefault("Remarks", None)
        r.setdefault("Month", "July")
        r.setdefault("First In", "09:00:00")
        r.setdefault("Last Out", "18:00:00")
        r.setdefault("Hrs", "08:00:00")
    return pd.DataFrame(rows)


class TestWOSpacing:
    def test_consecutive_wo_gap_seven_six_between(self):
        df = _base_df(
            [
                {"Empl Id": 1, "Attendance Date": "2025-06-23", "Duty Status": "Present"},
                {"Empl Id": 1, "Attendance Date": "2025-06-24", "Duty Status": "Present"},
                {"Empl Id": 1, "Attendance Date": "2025-06-25", "Duty Status": "Present"},
                {"Empl Id": 1, "Attendance Date": "2025-06-26", "Duty Status": "Present"},
                {"Empl Id": 1, "Attendance Date": "2025-06-27", "Duty Status": "Present"},
                {"Empl Id": 1, "Attendance Date": "2025-06-28", "Duty Status": "Present"},
                {"Empl Id": 1, "Attendance Date": "2025-06-29", "Duty Status": "WO"},
                {"Empl Id": 1, "Attendance Date": "2025-06-30", "Duty Status": "Present"},
                {"Empl Id": 1, "Attendance Date": "2025-07-01", "Duty Status": "Present"},
                {"Empl Id": 1, "Attendance Date": "2025-07-02", "Duty Status": "Present"},
                {"Empl Id": 1, "Attendance Date": "2025-07-03", "Duty Status": "Present"},
                {"Empl Id": 1, "Attendance Date": "2025-07-04", "Duty Status": "Present"},
                {"Empl Id": 1, "Attendance Date": "2025-07-05", "Duty Status": "Present"},
                {"Empl Id": 1, "Attendance Date": "2025-07-06", "Duty Status": "WO"},
            ]
        )
        df = build_wo_columns(df)
        assert validate_wo_output(df) == []
        wo = df[df["New_WO"] == "WO"].reset_index(drop=True)
        assert len(wo) == 2
        pos = df.index[df["New_WO"] == "WO"].tolist()
        assert pos[1] - pos[0] == 7
        assert df["week_day"].between(0, 6).all()
        assert (df.loc[df["New_WO"] == "WO", "week_day"] == 0).all()

    def test_no_wo_source_mode_shift(self):
        df = _base_df(
            [
                {"Empl Id": 1, "Attendance Date": "2025-06-23", "Duty Status": "Present"},
                {"Empl Id": 1, "Attendance Date": "2025-06-24", "Duty Status": "Present"},
            ]
        )
        df = build_wo_columns(df)
        assert (df["New_WO"] == "AS").all()
        assert (df["week_number"] == 0).all()
        assert validate_wo_output(df) == []

    def test_two_employees_independent(self):
        df = _base_df(
            [
                {"Empl Id": 1, "Attendance Date": "2025-06-23", "Duty Status": "WO"},
                {"Empl Id": 1, "Attendance Date": "2025-06-24", "Duty Status": "Present"},
                {"Empl Id": 2, "Attendance Date": "2025-06-23", "Duty Status": "Present"},
                {"Empl Id": 2, "Attendance Date": "2025-06-24", "Duty Status": "WO"},
                {"Empl Id": 2, "Attendance Date": "2025-06-25", "Duty Status": "Present"},
                {"Empl Id": 2, "Attendance Date": "2025-06-26", "Duty Status": "Present"},
                {"Empl Id": 2, "Attendance Date": "2025-06-27", "Duty Status": "Present"},
                {"Empl Id": 2, "Attendance Date": "2025-06-28", "Duty Status": "Present"},
                {"Empl Id": 2, "Attendance Date": "2025-06-29", "Duty Status": "Present"},
                {"Empl Id": 2, "Attendance Date": "2025-06-30", "Duty Status": "Present"},
                {"Empl Id": 2, "Attendance Date": "2025-07-01", "Duty Status": "WO"},
            ]
        )
        df = build_wo_columns(df)
        assert validate_wo_output(df) == []

    def test_week_cycle_indices(self):
        df = _base_df(
            [
                {"Empl Id": 1, "Attendance Date": "2025-06-23", "Duty Status": "WO"},
            ]
        )
        df = build_wo_columns(df)
        assert df.loc[0, "week_day"] == 0
        assert df.loc[0, "week_number"] == 0
        assert df.loc[0, "New_WO"] == "WO"


class TestTiming:
    def test_same_shift_copies_original_times(self):
        df = _base_df(
            [
                {
                    "Empl Id": 1,
                    "Attendance Date": "2025-06-23",
                    "Duty Status": "Present",
                    "Shift": "AS",
                    "First In": "05:00:00",
                    "Last Out": "14:00:00",
                    "Hrs": "08:00:00",
                },
            ]
        )
        df = build_wo_columns(df)
        timings = [
            {
                "shift": "AS",
                "first_in": "99:99:00",
                "last_out": "99:99:00",
                "hrs": "09:00:00",
            }
        ]
        df = add_new_timing_columns(df, timings)
        assert df.loc[0, "new_First In"] == "05:00:00"
        assert df.loc[0, "new_Last Out"] == "14:00:00"
        assert df.loc[0, "new_Hrs"] == "08:00:00"

    def test_different_shift_uses_sheet2(self):
        """Mode BS (4 BS vs 2 AS); row0 AS -> New_WO BS -> replace times from Sheet2."""
        rows = []
        shifts = ["AS", "BS", "BS", "BS", "BS", "AS"]
        for i, sh in enumerate(shifts):
            rows.append(
                {
                    "Empl Id": 1,
                    "Attendance Date": f"2025-06-{23+i:02d}",
                    "Duty Status": "Present",
                    "Shift": sh,
                    "First In": "05:00:00",
                    "Last Out": "14:00:00",
                    "Hrs": "08:00:00",
                }
            )
        rows.append(
            {
                "Empl Id": 1,
                "Attendance Date": "2025-06-29",
                "Duty Status": "WO",
                "Shift": "0",
                "First In": None,
                "Last Out": None,
                "Hrs": None,
            }
        )
        df = _base_df(rows)
        df = build_wo_columns(df)
        assert df.loc[0, "New_WO"] == "BS"
        timings = [
            {
                "shift": "BS",
                "first_in": "13:00:00",
                "last_out": "22:00:00",
                "hrs": "08:00:00",
            }
        ]
        df = add_new_timing_columns(df, timings)
        assert df.loc[0, "new_First In"] == "13:00:00"
        assert df.loc[0, "new_Last Out"] == "22:00:00"


class TestIntegration:
    @pytest.mark.skipif(
        not os.path.isfile(
            os.path.join(os.path.dirname(__file__), "Audit Work March 50_wo_final.xlsx")
        ),
        reason="Run process_wo_columns.py first to generate xlsx",
    )
    def test_full_file_passes_validation(self):
        path = os.path.join(os.path.dirname(__file__), "Audit Work March 50_wo_final.xlsx")
        try:
            df = pd.read_excel(path, engine="calamine")
        except Exception:
            df = pd.read_excel(path, engine="openpyxl")
        errs = validate_wo_output(df)
        assert errs == [], errs
