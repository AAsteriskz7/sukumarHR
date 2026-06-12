#!/usr/bin/env python3
"""
Script to process Audit Work March 46.xlsx and add new columns for WO allocation.

Updated Rules:
1. Find the first existing WO for each employee
2. Count 6 working days (skip Shift="0" and Remarks="Left")
3. Assign WO on the 6th working day
4. All 6 days in the week get the same shift based on max frequency in that week
5. Apply shift timings from Sheet2 for the assigned shift
6. Create new columns - don't modify existing rows
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from collections import Counter
import os

# File paths
file_path = "/Users/smartmysql/Documents/sukumarHR/Audit Work March 46.xlsx"
output_path = "/Users/smartmysql/Documents/sukumarHR/Audit Work March 46_processed.xlsx"

print(f"Reading file: {file_path}")

# Read Sheet1 with the actual employee data
df = pd.read_excel(file_path, sheet_name='Sheet1', header=0)

# Read Sheet2 for shift timings
df_shifts = pd.read_excel(file_path, sheet_name='Sheet2', header=0)
print(f"\nSheet2 shape: {df_shifts.shape}")

# Extract shift timings from Sheet2
shift_timings = {}

# Parse each row in Sheet2
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

# Convert Attendance Date to datetime
df['Attendance Date'] = pd.to_datetime(df['Attendance Date'])

# Sort by Employee ID and Attendance Date
df = df.sort_values(['Empl Id', 'Attendance Date']).reset_index(drop=True)

print(f"After sorting - shape: {df.shape}")

# Initialize new columns
df['Calculated_WO'] = 'N'
df['Week_Shift'] = ''  # Same shift for entire 6-day week
df['WO_Sequence'] = 0
df['Days_Since_Last_WO'] = 0

print("\nProcessing each employee...")

# Process each employee separately
employee_ids = df['Empl Id'].unique()
print(f"Total employees: {len(employee_ids)}")

for idx, emp_id in enumerate(employee_ids):
    if idx % 100 == 0:
        print(f"Processing employee {idx+1}/{len(employee_ids)}: {emp_id}")
    
    # Get employee data
    emp_mask = df['Empl Id'] == emp_id
    emp_indices = df[emp_mask].index.tolist()
    
    if len(emp_indices) == 0:
        continue
    
    # Find the first WO for this employee
    first_wo_idx = None
    for i, row_idx in enumerate(emp_indices):
        duty_status = df.loc[row_idx, 'Duty Status']
        if pd.notna(duty_status) and str(duty_status).strip().upper() == 'WO':
            first_wo_idx = i
            break
    
    # If no WO found, use first record as starting point
    if first_wo_idx is None:
        first_wo_idx = 0
    
    # Mark the first WO
    if first_wo_idx < len(emp_indices):
        df.loc[emp_indices[first_wo_idx], 'Calculated_WO'] = 'Y'
        df.loc[emp_indices[first_wo_idx], 'Week_Shift'] = 'WO'
        df.loc[emp_indices[first_wo_idx], 'WO_Sequence'] = 1
        df.loc[emp_indices[first_wo_idx], 'Days_Since_Last_WO'] = 0
    
    # Process in 6-day working cycles starting from after the first WO
    current_week = 1
    working_days_count = 0  # Count of actual working days (not skipped)
    week_shifts = []  # Track shifts in current week for max calculation
    week_indices = []  # Track indices in current week (only working days)
    days_since_wo = 0  # Count all days including skipped ones
    
    # Start from after the first WO
    start_idx = first_wo_idx + 1
    
    for i in range(start_idx, len(emp_indices)):
        row_idx = emp_indices[i]
        shift = df.loc[row_idx, 'Shift']
        remarks = df.loc[row_idx, 'Remarks']
        duty_status = df.loc[row_idx, 'Duty Status']
        
        # Check if this is already marked as WO in source data
        is_source_wo = pd.notna(duty_status) and str(duty_status).strip().upper() == 'WO'
        
        # Check if this record should be skipped (Shift="0" or Remarks="Left")
        is_skipped = False
        if pd.notna(shift) and str(shift) == '0':
            is_skipped = True
        if pd.notna(remarks) and str(remarks).strip().lower() == 'left':
            is_skipped = True
        
        # Always increment days_since_wo for tracking
        days_since_wo += 1
        
        if is_skipped:
            # Mark as skipped but preserve original values
            df.loc[row_idx, 'Calculated_WO'] = 'N'
            df.loc[row_idx, 'Week_Shift'] = str(shift) if pd.notna(shift) else ''
            df.loc[row_idx, 'WO_Sequence'] = current_week
            df.loc[row_idx, 'Days_Since_Last_WO'] = days_since_wo
            continue
        
        # Check if we need to assign WO (every 6 working days)
        if working_days_count >= 6 or is_source_wo:
            # This is a WO day - end of current week
            
            # Assign the most common shift to the entire previous week
            if len(week_shifts) > 0 and len(week_indices) > 0:
                shift_counts = Counter(week_shifts)
                most_common_shift = shift_counts.most_common(1)[0][0]
                
                # Update all working days in the week with the same shift
                for w_idx in week_indices:
                    df.loc[w_idx, 'Week_Shift'] = most_common_shift
            
            # Mark this as WO
            df.loc[row_idx, 'Calculated_WO'] = 'Y'
            df.loc[row_idx, 'Week_Shift'] = 'WO'
            df.loc[row_idx, 'WO_Sequence'] = current_week + 1
            df.loc[row_idx, 'Days_Since_Last_WO'] = 0
            
            # Reset for next week
            current_week += 1
            working_days_count = 0
            week_shifts = []
            week_indices = []
            days_since_wo = 0
        else:
            # Regular working day in the week
            df.loc[row_idx, 'Calculated_WO'] = 'N'
            df.loc[row_idx, 'WO_Sequence'] = current_week
            df.loc[row_idx, 'Days_Since_Last_WO'] = days_since_wo
            
            # Track shift for max calculation
            if pd.notna(shift) and str(shift) not in ['0', 'WO'] and str(shift) != 'nan':
                week_shifts.append(str(shift))
            week_indices.append(row_idx)
            
            working_days_count += 1
    
    # Handle the last week - assign most common shift to remaining working days
    if len(week_shifts) > 0 and len(week_indices) > 0:
        shift_counts = Counter(week_shifts)
        most_common_shift = shift_counts.most_common(1)[0][0]
        
        for w_idx in week_indices:
            if df.loc[w_idx, 'Calculated_WO'] != 'Y':
                df.loc[w_idx, 'Week_Shift'] = most_common_shift

print("\nProcessing complete!")

# Now update First In, Last Out, Hrs based on Week_Shift from Sheet2
print("\nUpdating shift timings from Sheet2...")

updated_count = 0
for idx, row in df.iterrows():
    week_shift = row['Week_Shift']
    if week_shift and week_shift in shift_timings and week_shift not in ['WO', '0']:
        timings = shift_timings[week_shift]
        df.loc[idx, 'First In'] = timings['First In']
        df.loc[idx, 'Last Out'] = timings['Last Out']
        df.loc[idx, 'Hrs'] = timings['Hrs']
        updated_count += 1

print(f"Updated {updated_count} records with shift timings from Sheet2")

# Display summary
print("\n=== SUMMARY ===")
print(f"Total records: {len(df)}")
print(f"WO records assigned: {(df['Calculated_WO'] == 'Y').sum()}")
print(f"Unique WO sequences: {df['WO_Sequence'].nunique()}")

# Show shift distribution in Week_Shift
print(f"\nWeek_Shift distribution:\n{df['Week_Shift'].value_counts()}")

# Show sample of results for first employee
print("\n=== SAMPLE OUTPUT (First 30 records of first employee) ===")
sample_cols = ['Empl Id', 'Attendance Date', 'Shift', 'Remarks', 'Duty Status', 
               'Calculated_WO', 'Week_Shift', 'First In', 'Last Out', 'Hrs', 
               'WO_Sequence', 'Days_Since_Last_WO']

first_emp = df['Empl Id'].iloc[0]
print(df[df['Empl Id'] == first_emp][sample_cols].head(30).to_string())

# Show sample where WO is assigned
wo_assigned = df[df['Calculated_WO'] == 'Y']
if len(wo_assigned) > 0:
    print(f"\n=== SAMPLE WO ASSIGNMENTS (First 10) ===")
    print(wo_assigned[sample_cols].head(10).to_string())

# Save to new Excel file
print(f"\nSaving to: {output_path}")
df.to_excel(output_path, index=False, sheet_name='Processed_Data')
print("Done!")
