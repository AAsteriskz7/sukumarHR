#!/usr/bin/env python3
"""Process test_50.xlsx with WO allocation logic."""

import pandas as pd
import numpy as np
from collections import Counter

# File paths
file_path = "/Users/smartmysql/Documents/sukumarHR/test_50.xlsx"
output_path = "/Users/smartmysql/Documents/sukumarHR/test_50_processed.xlsx"

print(f"Reading test file: {file_path}")

# Read Sheet "23" with employee data (different from 46.xlsx)
df = pd.read_excel(file_path, sheet_name='23', header=0)

# Read Sheet2 for shift timings
df_shifts = pd.read_excel(file_path, sheet_name='Sheet2', header=None)
print(f"\nSheet2 shape: {df_shifts.shape}")

# Extract shift timings from Sheet2
shift_timings = {}

for row_idx in range(df_shifts.shape[0]):
    row = df_shifts.iloc[row_idx]
    col_idx = 0
    while col_idx < len(row):
        val = row.iloc[col_idx]
        if pd.notna(val):
            shift_code = str(val).strip()
            if shift_code in ['GS', 'AS', 'CS', 'BS']:
                if col_idx + 3 < len(row):
                    first_in = row.iloc[col_idx + 1]
                    last_out = row.iloc[col_idx + 2]
                    hrs = row.iloc[col_idx + 3]
                    
                    if pd.notna(first_in) and pd.notna(last_out) and pd.notna(hrs):
                        try:
                            hrs_str = str(hrs)
                            if ':' in hrs_str:
                                parts = hrs_str.split(':')
                                hrs_val = int(parts[0]) + int(parts[1])/60
                            else:
                                hrs_val = 8.5
                            
                            if hrs_val > 8 and shift_code not in shift_timings:
                                shift_timings[shift_code] = {
                                    'First In': first_in,
                                    'Last Out': last_out,
                                    'Hrs': hrs
                                }
                        except:
                            pass
        col_idx += 1

print(f"Extracted shift timings: {shift_timings}")

print(f"\nOriginal shape: {df.shape}")
print(f"Original columns: {df.columns.tolist()}")
print("\nOriginal sample:")
print(df.head(10).to_string())

# Convert Attendance Date to datetime
df['Attendance Date'] = pd.to_datetime(df['Attendance Date'])

# Sort by Employee ID and Attendance Date
df = df.sort_values(['Empl Id', 'Attendance Date']).reset_index(drop=True)

print(f"\nAfter sorting - shape: {df.shape}")

# Initialize new columns
df['Calculated_WO'] = 'N'
df['Week_Shift'] = ''
df['WO_Sequence'] = 0
df['Days_Since_Last_WO'] = 0

print("\n" + "="*80)
print("PROCESSING EACH EMPLOYEE")
print("="*80)

employee_ids = df['Empl Id'].unique()
print(f"Total employees: {len(employee_ids)}: {list(employee_ids)}")

for idx, emp_id in enumerate(employee_ids):
    print(f"\n{'='*80}")
    print(f"EMPLOYEE {emp_id} ({idx+1}/{len(employee_ids)})")
    print('='*80)
    
    emp_mask = df['Empl Id'] == emp_id
    emp_indices = df[emp_mask].index.tolist()
    
    if len(emp_indices) == 0:
        continue
    
    print(f"Total records: {len(emp_indices)}")
    
    # Find first WO
    first_wo_idx = None
    for i, row_idx in enumerate(emp_indices):
        duty_status = df.loc[row_idx, 'Duty Status']
        if pd.notna(duty_status) and str(duty_status).strip().upper() == 'WO':
            first_wo_idx = i
            print(f"First WO at index {i} (Date: {df.loc[row_idx, 'Attendance Date'].strftime('%Y-%m-%d')})")
            break
    
    if first_wo_idx is None:
        first_wo_idx = 0
        print(f"No WO found. Starting from index 0")
    
    # Mark first WO
    if first_wo_idx < len(emp_indices):
        df.loc[emp_indices[first_wo_idx], 'Calculated_WO'] = 'Y'
        df.loc[emp_indices[first_wo_idx], 'Week_Shift'] = 'WO'
        df.loc[emp_indices[first_wo_idx], 'WO_Sequence'] = 1
        df.loc[emp_indices[first_wo_idx], 'Days_Since_Last_WO'] = 0
    
    # Process 6-day cycles
    current_week = 1
    working_days_count = 0
    week_shifts = []
    week_indices = []
    days_since_wo = 0
    
    start_idx = first_wo_idx + 1
    
    for i in range(start_idx, len(emp_indices)):
        row_idx = emp_indices[i]
        shift = df.loc[row_idx, 'Shift']
        remarks = df.loc[row_idx, 'Remarks']
        duty_status = df.loc[row_idx, 'Duty Status']
        date = df.loc[row_idx, 'Attendance Date']
        
        is_source_wo = pd.notna(duty_status) and str(duty_status).strip().upper() == 'WO'
        
        # Skip conditions
        is_skipped = False
        skip_reason = ""
        if pd.notna(shift) and str(shift) == '0':
            is_skipped = True
            skip_reason = "Shift=0"
        if pd.notna(remarks) and str(remarks).strip().lower() == 'left':
            is_skipped = True
            skip_reason = "Remarks=Left"
        
        days_since_wo += 1
        
        if is_skipped:
            df.loc[row_idx, 'Calculated_WO'] = 'N'
            df.loc[row_idx, 'Week_Shift'] = str(shift) if pd.notna(shift) else ''
            df.loc[row_idx, 'WO_Sequence'] = current_week
            df.loc[row_idx, 'Days_Since_Last_WO'] = days_since_wo
            print(f"  {date.strftime('%Y-%m-%d')}: SKIP ({skip_reason})")
            continue
        
        # Check for WO assignment (every 6 working days)
        if working_days_count >= 6 or is_source_wo:
            # Assign most common shift to week
            if len(week_shifts) > 0 and len(week_indices) > 0:
                shift_counts = Counter(week_shifts)
                most_common_shift = shift_counts.most_common(1)[0][0]
                print(f"  Week {current_week}: {dict(shift_counts)} -> '{most_common_shift}' for {len(week_indices)} days")
                
                for w_idx in week_indices:
                    df.loc[w_idx, 'Week_Shift'] = most_common_shift
            
            # Mark WO
            df.loc[row_idx, 'Calculated_WO'] = 'Y'
            df.loc[row_idx, 'Week_Shift'] = 'WO'
            df.loc[row_idx, 'WO_Sequence'] = current_week + 1
            df.loc[row_idx, 'Days_Since_Last_WO'] = 0
            print(f"  {date.strftime('%Y-%m-%d')}: WO ASSIGNED (day {working_days_count})")
            
            # Reset
            current_week += 1
            working_days_count = 0
            week_shifts = []
            week_indices = []
            days_since_wo = 0
        else:
            # Regular working day
            df.loc[row_idx, 'Calculated_WO'] = 'N'
            df.loc[row_idx, 'WO_Sequence'] = current_week
            df.loc[row_idx, 'Days_Since_Last_WO'] = days_since_wo
            
            if pd.notna(shift) and str(shift) not in ['0', 'WO'] and str(shift) != 'nan':
                week_shifts.append(str(shift))
            week_indices.append(row_idx)
            
            working_days_count += 1
    
    # Handle last week
    if len(week_shifts) > 0 and len(week_indices) > 0:
        shift_counts = Counter(week_shifts)
        most_common_shift = shift_counts.most_common(1)[0][0]
        print(f"  Week {current_week} (final): {dict(shift_counts)} -> '{most_common_shift}' for {len(week_indices)} days")
        
        for w_idx in week_indices:
            if df.loc[w_idx, 'Calculated_WO'] != 'Y':
                df.loc[w_idx, 'Week_Shift'] = most_common_shift

print("\n" + "="*80)
print("UPDATING SHIFT TIMINGS FROM SHEET2")
print("="*80)

updated_count = 0
for idx, row in df.iterrows():
    week_shift = row['Week_Shift']
    if week_shift and week_shift in shift_timings and week_shift not in ['WO', '0']:
        timings = shift_timings[week_shift]
        df.loc[idx, 'First In'] = timings['First In']
        df.loc[idx, 'Last Out'] = timings['Last Out']
        df.loc[idx, 'Hrs'] = timings['Hrs']
        updated_count += 1

print(f"Updated {updated_count} records with shift timings")

# Summary
print("\n" + "="*80)
print("FINAL SUMMARY")
print("="*80)
print(f"Total records: {len(df)}")
print(f"WO records: {(df['Calculated_WO'] == 'Y').sum()}")
print(f"Week_Shift distribution:\n{df['Week_Shift'].value_counts()}")

# Show detailed results
print("\n" + "="*80)
print("DETAILED RESULTS")
print("="*80)

sample_cols = ['Empl Id', 'Attendance Date', 'Shift', 'Remarks', 'Duty Status', 
               'Calculated_WO', 'Week_Shift', 'First In', 'Last Out', 'Hrs', 
               'WO_Sequence', 'Days_Since_Last_WO']

for emp_id in employee_ids:
    print(f"\n{'='*60}")
    print(f"EMPLOYEE {emp_id}")
    print('='*60)
    emp_data = df[df['Empl Id'] == emp_id][sample_cols].head(40)
    print(emp_data.to_string())

print(f"\n{'='*80}")
print(f"Saving to: {output_path}")
df.to_excel(output_path, index=False, sheet_name='Processed_Data')
print("DONE!")
