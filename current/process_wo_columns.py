#!/usr/bin/env python3
"""
Add New_WO, week_day, week_number, and new_First In / new_Last Out / new_Hrs.

WO rhythm (unchanged):
  - Per Empl Id, sort by Attendance Date.
  - If no Duty Status == WO exists: no WO cycle; week_day/week_number default 0; New_WO uses row rules without weekly mode.
  - Else first WO at sorted index i0; calculated WO on rows i0, i0+7, i0+14, ...

New_WO column (not Y/N):
  - Calculated WO day -> "WO"
  - Shift "0" or Remarks "Left" -> keep original Shift value (no change)
  - Source WO on a non-calculated-WO row, or leave/absence-type Duty Status -> "Absent"
  - Otherwise -> most frequent Shift in that 7-row week (ties: first mode from Counter)

week_day: 0..6 = (row_index - i0) % 7 after first WO; before i0 use pos % 7; no-WO employees use pos % 7.

week_number: per employee, 0,1,2,... = (row_index - i0) // 7 when pos >= i0; before i0 use 0; no-WO employees 0.

Timing (new_First In, new_Last Out, new_Hrs):
  - Start as copies of First In / Last Out / Hrs.
  - If Shift==0 or Remarks Left: unchanged.
  - If New_WO is WO or Absent: unchanged.
  - If original Shift matches assigned New_WO shift code: unchanged.
  - Else: random First In / Last Out / Hrs from Sheet2 for that assigned shift (GS/AS/CS/BS).

Usage:
  python3 process_wo_columns.py
  python3 process_wo_columns.py --validate "Audit Work March 50_wo_final.xlsx"

  Timestamped output (avoids overwriting), e.g. Auro attendance .xls:
  python3 process_wo_columns.py \\
    --input "Auro_Attendance_rept_AFD_U-XV_2025_12.xls" \\
    --sheet "Auro_Attendance_rept_AFD_U-XV_2" \\
    --sheet2-workbook "Audit Work March 50.xlsx" \\
    --output "Auro_Attendance_rept_AFD_U-XV_2025_12_wo_final.xlsx" \\
    --output-timestamp
  # Sheet2 timings from Audit file; attendance from --input. Writes e.g. ..._wo_final_20250323_143052.xlsx
"""

from __future__ import annotations

import argparse
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


# Max allowed total hours (in minutes) per shift category
_SHIFT_MAX_MINS: dict[str, int] = {
    "AS": 8 * 60 + 10,   # 8h 10m
    "BS": 8 * 60 + 10,   # 8h 10m
    "CS": 8 * 60 + 10,   # 8h 10m
    "GS": 8 * 60 + 35,   # 8h 35m
}


def generate_synthetic_timing(shift_code: str) -> dict[str, str]:
    """Generate realistic In/Out/Hrs punches for a given shift.

    Start time varies randomly up to 10 min early from base start.
    Total hours are capped at 8:10 for AS/BS/CS and 8:35 for GS.
    Extra minutes (2..cap_headroom) are split between early arrival and late departure.
    """
    sc = shift_code.upper()

    # Base times in minutes past midnight (start, end)
    base_times = {
        "AS": (6 * 60, 14 * 60),            # 6:00 AM to 2:00 PM (8h)
        "BS": (14 * 60, 22 * 60),           # 2:00 PM to 10:00 PM (8h)
        "CS": (22 * 60, 6 * 60 + 24 * 60),  # 10:00 PM to 6:00 AM (8h)
        "GS": (9 * 60, 17 * 60 + 30),       # 9:00 AM to 5:30 PM (8h 30m)
    }

    if sc not in base_times:
        return {}

    base_start, base_end = base_times[sc]
    base_duration = base_end - base_start
    max_duration = _SHIFT_MAX_MINS[sc]
    max_extra = max_duration - base_duration  # headroom before hitting the cap

    # Extra minutes are at least 2 and at most the cap headroom (e.g. 10 for ABC, 5 for GS)
    extra_mins = random.randint(2, max(2, max_extra))
    extra_mins = min(extra_mins, max_extra)  # hard cap

    # Randomly distribute between arriving early and leaving late
    # Start can be up to 10 min early (but not more than extra_mins)
    early_mins = random.randint(0, min(extra_mins, 10))
    late_mins = extra_mins - early_mins

    in_punch_mins = base_start - early_mins
    out_punch_mins = base_end + late_mins
    duration_mins = base_duration + extra_mins  # guaranteed <= max_duration

    def format_time(mins: int) -> str:
        m = mins % (24 * 60)
        hours = m // 60
        minutes = m % 60
        return f"{hours:02d}:{minutes:02d}:00"

    return {
        "first_in": format_time(in_punch_mins),
        "last_out": format_time(out_punch_mins),
        "hrs": format_time(duration_mins),
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
        nw = df.at[i, "New_WO"]

        s_nw = str(nw).strip() if pd.notna(nw) else ""
        if not s_nw or s_nw.upper() in ("WO", "ABSENT"):
            continue

        assigned = s_nw.upper()

        # Always apply capped synthetic timing for any valid shift assignment
        if assigned in ("AS", "BS", "CS", "GS"):
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


def is_compound_duty(duty: object) -> bool:
    """Return True if Duty Status is a slash-compound value like SL/P, P/HD, CL/P, etc.

    These rows should keep their original Shift in New_WO (not marked Absent,
    not reassigned to the weekly mode shift).
    """
    if pd.isna(duty):
        return False
    return "/" in str(duty).strip()


def is_absence_display(
    duty: object, *, is_calculated_wo: bool, is_legacy_wo: bool
) -> bool:
    """Non-Present leave types -> show 'Absent' in New_WO (unless skip/cal WO/compound duty)."""
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
    # Slash-compound statuses (e.g. SL/P, P/HD) are handled separately
    if "/" in u:
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

                # Slash-compound Duty Status (e.g. SL/P, P/HD, CL/P) -> keep original Shift
                if is_compound_duty(duty):
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


def _output_path_with_timestamp(path: str) -> str:
    """Insert _YYYYMMDD_HHMMSS before the file extension."""
    p = Path(path)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    return str(p.parent / f"{p.stem}_{ts}{p.suffix}")


def _write_xlsx(df: pd.DataFrame, path: str, *, sheet_name: str = "Processed_Data") -> None:
    try:
        df.to_excel(path, index=False, sheet_name=sheet_name, engine="xlsxwriter")
    except ImportError:
        df.to_excel(path, index=False, sheet_name=sheet_name, engine="openpyxl")


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


def main() -> int:
    parser = argparse.ArgumentParser(description="Add New_WO, week_day, week_number columns.")
    parser.add_argument(
        "--input",
        default="/Users/smartmysql/Documents/sukumarHR/Audit Work March 50.xlsx",
        help="Input Excel path",
    )
    parser.add_argument(
        "--output",
        default="/Users/smartmysql/Documents/sukumarHR/Audit Work March 50_wo_final.xlsx",
        help="Output path (.xlsx default; .csv/.parquet for very large files)",
    )
    parser.add_argument(
        "--output-timestamp",
        action="store_true",
        help="Append _YYYYMMDD_HHMMSS before extension so each run writes a new file",
    )
    parser.add_argument("--sheet", default="23", help="Worksheet name")
    parser.add_argument(
        "--seed",
        type=int,
        default=None,
        help="Random seed for timing picks (default: non-deterministic)",
    )
    parser.add_argument(
        "--validate",
        metavar="FILE",
        help="Run validation on output (csv, parquet, or xlsx) and exit (0=ok, 1=failed)",
    )
    args = parser.parse_args()

    if args.validate:
        try:
            path = args.validate
            if path.lower().endswith((".parquet", ".pq")):
                df = pd.read_parquet(path)
            elif path.lower().endswith((".xlsx", ".xls")):
                try:
                    df = pd.read_excel(path, engine="calamine")
                except Exception:
                    df = pd.read_excel(path, engine="openpyxl")
            else:
                df = pd.read_csv(path)
        except Exception as e:
            print(f"ERROR: could not read {args.validate!r}: {e}", file=sys.stderr)
            return 1
        errs = validate_wo_output(df)
        if errs:
            print(f"Validation FAILED ({len(errs)} issue(s)):")
            for e in errs[:200]:
                print(f"  - {e}")
            if len(errs) > 200:
                print(f"  ... and {len(errs) - 200} more")
            return 1
        print("Validation OK: WO spacing, week_day 0-6, week_number per employee.")
        return 0

    t0 = time.perf_counter()
    print(f"Reading: {args.input} (sheet={args.sheet!r})")
    df = _read_excel(args.input, args.sheet)
    t1 = time.perf_counter()
    print(f"  Rows: {len(df):,}  ({t1 - t0:.1f}s)")

    if "Hrs" not in df.columns and "Total Hours Worked" in df.columns:
        df = df.rename(columns={"Total Hours Worked": "Hrs"})

    if "Empl Id" not in df.columns or "Attendance Date" not in df.columns:
        print("ERROR: required columns Empl Id, Attendance Date not found.", file=sys.stderr)
        return 1
    if "Duty Status" not in df.columns:
        print("ERROR: required column Duty Status not found.", file=sys.stderr)
        return 1
    if "Remarks" not in df.columns:
        df["Remarks"] = ""
    if "Shift" not in df.columns:
        print("ERROR: required column Shift not found.", file=sys.stderr)
        return 1
    for c in ("First In", "Last Out", "Hrs"):
        if c not in df.columns:
            print(f"ERROR: required column {c!r} not found.", file=sys.stderr)
            return 1

    if args.seed is not None:
        random.seed(args.seed)

    df = build_wo_columns(df)
    df = add_new_timing_columns(df)

    out = args.output
    if args.output_timestamp:
        out = _output_path_with_timestamp(out)
    t2 = time.perf_counter()
    if out.lower().endswith(".csv"):
        df.to_csv(out, index=False)
    elif out.lower().endswith((".parquet", ".pq")):
        try:
            df.to_parquet(out, index=False)
        except ImportError:
            print(
                "ERROR: pip install pyarrow (or fastparquet) for Parquet output.",
                file=sys.stderr,
            )
            return 1
    else:
        _write_xlsx(df, out)

    t3 = time.perf_counter()
    print(f"Written: {out}")
    print(f"  Compute: {t2 - t1:.1f}s  Write: {t3 - t2:.1f}s  Total: {t3 - t0:.1f}s")
    wo_count = (df["New_WO"].astype(str) == "WO").sum()
    print(f"  New_WO=WO rows: {wo_count:,}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
