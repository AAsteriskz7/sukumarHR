#!/usr/bin/env python3
"""
Add New_WO, week_day, week_number, and new_First In / new_Last Out / new_Hrs.

This is a copy of the core processing module placed alongside app.py so that
the Streamlit Cloud deployment can import it without path hacks.
"""

from __future__ import annotations

import random
import sys
import time
from collections import Counter
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd


def _read_excel(path: str, sheet: str) -> pd.DataFrame:
    try:
        return pd.read_excel(path, sheet_name=sheet, engine="calamine")
    except Exception:
        return pd.read_excel(path, sheet_name=sheet, engine="openpyxl")


def generate_synthetic_timing(shift_code: str) -> dict[str, str]:
    """Generate realistic In/Out/Hrs punches for a given shift with variance matching user examples."""
    sc = shift_code.upper()

    # Base times in minutes past midnight (start, end)
    base_times = {
        "AS": (6 * 60, 14 * 60),             # 6:00 AM to 2:00 PM (8h)
        "BS": (14 * 60, 22 * 60),            # 2:00 PM to 10:00 PM (8h)
        "CS": (22 * 60, 6 * 60 + 24 * 60),   # 10:00 PM to 6:00 AM (8h)
        "GS": (9 * 60, 17 * 60 + 30),        # 9:00 AM to 5:30 PM (8h 30m)
    }

    if sc not in base_times:
        return {}

    base_start, base_end = base_times[sc]
    base_duration = base_end - base_start

    # Generate 2 to 10 extra minutes above the base duration
    extra_mins = random.randint(2, 10)

    # Randomly distribute the extra minutes between arriving early and leaving late
    early_mins = random.randint(0, extra_mins)
    late_mins = extra_mins - early_mins

    in_punch_mins = base_start - early_mins
    out_punch_mins = base_end + late_mins
    duration_mins = base_duration + extra_mins

    def format_time(mins: int) -> str:
        # Normalize to 24 hour clock string formatting
        m = mins % (24 * 60)
        hours = m // 60
        minutes = m % 60
        return f"{hours:02d}:{minutes:02d}:00"

    return {
        "first_in": format_time(in_punch_mins),
        "last_out": format_time(out_punch_mins),
        "hrs": format_time(duration_mins)
    }


def add_new_timing_columns(
    df: pd.DataFrame,
    *,
    first_in_col: str = "First In",
    last_out_col: str = "Last Out",
    hrs_col: str = "Hrs",
    new_first: str = "new_First In",
    new_last: str = "new_Last Out",
    new_hrs: str = "new_Hrs",
) -> pd.DataFrame:
    """
    Enforces synthetic times for assigned regular shifts (GS/AS/BS/CS) with maximum hour rules.
    Skips Shift==0, Remarks Left, WO, Absent.
    """
    for c in (first_in_col, last_out_col, hrs_col):
        if c not in df.columns:
            raise ValueError(f"Missing column {c!r}")

    nfi = df[first_in_col].copy()
    nlo = df[last_out_col].copy()
    nh = df[hrs_col].copy()

    for i in range(len(df)):
        shift = df.at[i, "Shift"]
        nw = df.at[i, "New_WO"]

        s_nw = str(nw).strip() if pd.notna(nw) else ""
        if not s_nw or s_nw.upper() in ("WO", "ABSENT"):
            continue

        assigned = s_nw.upper()
        orig_shift = str(shift).strip().upper() if pd.notna(shift) else ""

        # Apply synthetic times only if assigned a main shift AND it differs from the original
        if assigned in ("AS", "BS", "CS", "GS") and orig_shift != assigned:
            syn = generate_synthetic_timing(assigned)
            if syn:
                nfi.iloc[i] = syn["first_in"]
                nlo.iloc[i] = syn["last_out"]
                nh.iloc[i] = syn["hrs"]

    df[new_first] = nfi
    df[new_last] = nlo
    df[new_hrs] = nh
    return df


def _first_wo_position_within_group(duty: pd.Series) -> float:
    m = duty.astype(str).str.strip().str.upper() == "WO"
    if not m.any():
        return np.nan
    return float(np.argmax(m.to_numpy()))


def should_skip_shift(shift: object, remarks: object, duty: object = None) -> bool:
    if pd.notna(shift):
        s_str = str(shift).strip().upper()
        if s_str == "0" or s_str in ("HO", "CL", "SL", "EL", "OD"):
            return True
    if pd.notna(remarks):
        r_str = str(remarks).strip().upper()
        if r_str == "LEFT" or r_str in ("HO", "CL", "SL", "EL", "OD"):
            return True
    if pd.notna(duty):
        d_str = str(duty).strip().upper()
        if d_str in ("HO", "CL", "SL", "EL", "OD"):
            return True
    return False


def is_present_like(duty: object) -> bool:
    """Present / on-duty: use weekly mode shift."""
    if pd.isna(duty):
        return False
    u = str(duty).strip().upper()
    return u in ("PRESENT", "OD")


def is_source_wo(duty: object, shift: object = None) -> bool:
    if pd.notna(duty) and str(duty).strip().upper() == "WO":
        return True
    if pd.notna(shift) and str(shift).strip().upper() == "WO":
        return True
    return False


def is_absence_display(
    duty: object, *, is_calculated_wo: bool, is_legacy_wo: bool
) -> bool:
    """Non-Present leave types -> show 'Absent' in New_WO (unless skip/cal WO)."""
    if is_calculated_wo:
        return False
    if is_legacy_wo:
        return False
    if pd.isna(duty):
        return False
    u = str(duty).strip().upper()
    if u in ("PRESENT", "OD"):
        return False
    if u == "WO":
        return False
    return True


def _normalize_shift(val: object) -> str | None:
    if pd.isna(val):
        return None
    s = str(val).strip()
    if not s or s.upper() == "NAN":
        return None
    return s


def _mode_shift(indices: list[int], df: pd.DataFrame, cal_wo: np.ndarray) -> str | None:
    shifts: list[str] = []
    for i in indices:
        duty = df.at[i, "Duty Status"]
        shift = df.at[i, "Shift"]
        remarks = df.at[i, "Remarks"]
        if cal_wo[i]:
            continue
        if should_skip_shift(shift, remarks, duty):
            continue
        if is_present_like(duty):
            sh = _normalize_shift(shift)
            if sh and sh.upper() not in ("WO", "ES"):
                shifts.append(sh)
    if not shifts:
        return None
    return Counter(shifts).most_common(1)[0][0]


def build_wo_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Add week_day, week_number, New_WO."""
    df = df.sort_values(["Empl Id", "Attendance Date"], kind="mergesort").reset_index(
        drop=True
    )
    n = len(df)

    # Detect WO from EITHER Duty Status or Shift column
    duty_is_wo = df["Duty Status"].astype(str).str.strip().str.upper() == "WO"
    shift_is_wo = df["Shift"].astype(str).str.strip().str.upper() == "WO"
    df["_wo_flag"] = (duty_is_wo | shift_is_wo).map({True: "WO", False: ""})

    g = df.groupby("Empl Id", sort=False)
    pos = g.cumcount().to_numpy()
    first_wo = g["_wo_flag"].transform(_first_wo_position_within_group).to_numpy()

    has = ~np.isnan(first_wo)
    diff = pos - first_wo
    cal_wo = has & (diff >= 0) & (np.mod(diff, 7) == 0)
    cal_wo_arr = cal_wo

    week_day = np.zeros(n, dtype=np.int64)
    week_num = np.zeros(n, dtype=np.int64)
    for i in range(n):
        fp = first_wo[i]
        p = int(pos[i])
        if np.isnan(fp):
            week_day[i] = p % 7
            week_num[i] = 0
        else:
            fpi = int(fp)
            if p < fpi:
                week_day[i] = p % 7
                week_num[i] = 0
            else:
                rel = p - fpi
                week_day[i] = rel % 7
                week_num[i] = rel // 7

    new_wo = np.empty(n, dtype=object)

    for _, grp in df.groupby("Empl Id", sort=False):
        idx_list = grp.index.tolist()
        i0 = first_wo[idx_list[0]]
        segments: list[list[int]] = []
        if np.isnan(i0):
            segments = [idx_list]
        else:
            i0i = int(i0)
            if i0i > 0:
                segments.append(idx_list[:i0i])
            rest = idx_list[i0i:]
            for s in range(0, len(rest), 7):
                segments.append(rest[s : s + 7])

        mode_cache: dict[tuple[int, int], str | None] = {}

        def seg_key(seg: list[int]) -> tuple[int, int]:
            return (seg[0], seg[-1]) if seg else (-1, -1)

        for seg in segments:
            if not seg:
                continue
            mode_cache[seg_key(seg)] = _mode_shift(seg, df, cal_wo_arr)

        for seg in segments:
            if not seg:
                continue
            mode = mode_cache[seg_key(seg)]
            for row_idx in seg:
                duty = df.at[row_idx, "Duty Status"]
                shift = df.at[row_idx, "Shift"]
                remarks = df.at[row_idx, "Remarks"]
                cwo = bool(cal_wo_arr[row_idx])
                legacy_wo = is_source_wo(duty, shift) and not cwo

                if cwo:
                    new_wo[row_idx] = "WO"
                    continue

                orig_normalized = _normalize_shift(shift)
                if orig_normalized and orig_normalized.upper() == "ES":
                    k = idx_list.index(row_idx)
                    replacement = None
                    # Look downward past consecutive ES/WO rows
                    for look in range(k + 1, len(idx_list)):
                        below_idx = idx_list[look]
                        below_sh = _normalize_shift(df.at[below_idx, "Shift"])
                        if below_sh and below_sh.upper() not in ("ES", "WO"):
                            replacement = below_sh
                            break
                    # Fallback: look upward past ES/WO rows
                    if not replacement:
                        for look in range(k - 1, -1, -1):
                            above_idx = idx_list[look]
                            above_sh = _normalize_shift(df.at[above_idx, "Shift"])
                            if above_sh and above_sh.upper() not in ("ES", "WO"):
                                replacement = above_sh
                                break
                    if replacement:
                        new_wo[row_idx] = replacement
                        continue

                if should_skip_shift(shift, remarks, duty):
                    new_wo[row_idx] = _normalize_shift(shift) or ""
                    continue

                if is_absence_display(duty, is_calculated_wo=cwo, is_legacy_wo=legacy_wo):
                    new_wo[row_idx] = "Absent"
                    continue
                if mode is None:
                    new_wo[row_idx] = _normalize_shift(shift) or ""
                else:
                    new_wo[row_idx] = mode

    df["week_day"] = week_day
    df["week_number"] = week_num
    df["New_WO"] = new_wo
    df = df.drop(columns=["_wo_flag"], errors="ignore")
    return df


def validate_wo_output(
    df: pd.DataFrame,
    *,
    date_col: str = "Attendance Date",
    emp_col: str = "Empl Id",
    wo_col: str = "New_WO",
    week_day_col: str = "week_day",
    week_num_col: str = "week_number",
) -> list[str]:
    errors: list[str] = []
    required = {date_col, emp_col, wo_col, week_day_col, week_num_col}
    missing = required - set(df.columns)
    if missing:
        return [f"Missing columns: {sorted(missing)}"]

    wd = pd.to_numeric(df[week_day_col], errors="coerce")
    bad_wd = wd.isna() | (wd < 0) | (wd > 6)
    if bad_wd.any():
        errors.append(f"{week_day_col} must be 0..6: {int(bad_wd.sum())} bad row(s)")

    for emp_id, sub in df.groupby(emp_col, sort=False):
        sub = sub.sort_values(date_col, kind="mergesort").reset_index(drop=True)
        wo_mask = sub[wo_col].astype(str) == "WO"
        y_ix = np.flatnonzero(wo_mask.to_numpy())
        if len(y_ix) < 2:
            continue
        for i in range(len(y_ix) - 1):
            a, b = int(y_ix[i]), int(y_ix[i + 1])
            gap = b - a
            between = sub.iloc[a + 1 : b]
            if gap != 7 or len(between) != 6:
                errors.append(
                    f"Empl Id {emp_id}: consecutive WO rows gap {gap}, "
                    f"between={len(between)} (expected 7 and 6)"
                )
    return errors
