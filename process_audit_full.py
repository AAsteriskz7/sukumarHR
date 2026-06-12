#!/usr/bin/env python3
"""
Script to process Audit Work March 46.xlsx and add new columns for WO allocation.

Updated Rules:
1. Find the first existing WO for each employee, then assign WO every 6 days after that
2. Each week (6 records between WOs) gets the same shift based on max frequency in that week
3. Skip records where Shift is "0" or Remarks is "Left"
4. Shift timings (First In, Last Out, Hrs) should come from Sheet2 where Hrs > 8 hours
5. Create new columns - don't modify existing rows

New columns to add:
- Calculated_WO: Y/N flag for WO allocation
- Calculated_Shift: Shift for the week (AS/GS/CS/BS) based on max frequency
- Week_Shift: Same shift applied to entire 6-day week
- WO_Sequence: Sequence number for tracking WO cycles
- Days_Since_Last_WO: Counter for days since last WO
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

# Read Sheet2 for shift timings (GS, AS, CS, BS with First In, Last Out, Hrs > 8)
df_shifts = pd.read_excel(file_path, sheet_name='Sheet2', header=None)
print(f"\nSheet2 shape: {df_shifts.shape}")
print(f"Sheet2 sample:\n{df_shifts.head(10).to_string()}")

# Extract shift timings from Sheet2
# Structure appears to be: GS, time1, time2, time3, empty, AS, time1, time2, time3, etc.
shift_timings = {}

# Parse Sheet2 to get shift timings
# Looking for pattern: Shift Code, First In, Last Out, Hrs
for col_idx in range(0, df_shifts.shape[1], 5):  # Every 5th column starts a new shift block
    if col_idx < df_shifts.shape[1]:
        shift_code = str(df_shifts.iloc[0, col_idx]).strip() if pd.notna(df_shifts.iloc[0, col_idx]) else None
        if shift_code and shift_code in ['GS', 'AS', 'CS', 'BS']:
            # Extract times from first row (could be different row with actual times)
            for row_idx in range(min(20, df_shifts.shape[0])):
                first_in = df_shifts.iloc[row_idx, col_idx + 1] if (col_idx + 1) < df_shifts.shape[1] else None
                last_out = df_shifts.iloc[row_idx, col_idx + 2] if (col_idx + 2) < df_shifts.shape[1] else None
                hrs = df_shifts.iloc[row_idx, col_idx + 3] if (col_idx + 3) < df_shifts.shape[1] else None
                
                # Check if this row has valid time values and Hrs > 8
                if pd.notna(first_in) and pd.notna(last_out) and pd.notna(hrs):
                    try:
                        # Parse hrs - could be time string or numeric
                        if isinstance(hrs, str) and ':' in hrs:
                            parts = hrs.split(':')
                            hrs_val = int(parts[0]) + int(parts[1])/60 if len(parts) >= 2 else 0
                        else:
                            hrs_val = float(hrs) if pd.notna(hrs) else 0
                        
                        if hrs_val > 8:
                            shift_timings[shift_code] = {
                                'First In': first_in,
                                'Last Out': last_out,
                                'Hrs': hrs
                            }
                            break
                    except:
                        pass

print(f"\nExtracted shift timings: {shift_timings}")

print(f"\nOriginal shape: {df.shape}")
print(f"Original columns: {df.columns.tolist()}")
print("\nSample data:")
print(df.head(10).to_string())

# Convert Attendance Date to datetime
df['Attendance Date'] = pd.to_datetime(df['Attendance Date'])

# Sort by Employee ID and Attendance Date
df = df.sort_values(['Empl Id', 'Attendance Date']).reset_index(drop=True)

print(f"\nAfter sorting - shape: {df.shape}")

# Initialize new columns with default values
df['Calculated_WO'] = 'N'
df['Week_Shift'] = ''  # Same shift for entire 6-day week
df['WO_Sequence'] = 0
df['Days_Since_Last_WO'] = 0
df['Week_Number'] = 0  # Track which week the record belongs to

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
    
    # If no WO found, skip this employee or use first record as starting point
    if first_wo_idx is None:
        first_wo_idx = 0
    
    # Mark the first WO
    if first_wo_idx < len(emp_indices):
        df.loc[emp_indices[first_wo_idx], 'Calculated_WO'] = 'Y'
        df.loc[emp_indices[first_wo_idx], 'Week_Shift'] = 'WO'
        df.loc[emp_indices[first_wo_idx], 'WO_Sequence'] = 1
        df.loc[emp_indices[first_wo_idx], 'Days_Since_Last_WO'] = 0
        df.loc[emp_indices[first_wo_idx], 'Week_Number'] = 1
    
    # Now process in 6-day cycles starting from after the first WO
    current_week = 1
    days_in_week = 0
    week_shifts = []  # Track shifts in current week for max calculation
    week_indices = []  # Track indices in current week
    
    # Start from after the first WO
    start_idx = first_wo_idx + 1
    
    for i in range(start_idx, len(emp_indices)):
        row_idx = emp_indices[i]
        shift = df.loc[row_idx, 'Shift']
        remarks = df.loc[row_idx, 'Remarks']
        duty_status = df.loc[row_idx, 'Duty Status']
        
        # Check if this is already marked as WO in source data
        is_source_wo = pd.notna(duty_status) and str(duty_status).strip().upper() == 'WO'
        
        # Skip if Shift is "0"
        if pd.notna(shift) and str(shift) == '0':
            df.loc[row_idx, 'Calculated_WO'] = 'N'
            df.loc[row_idx, 'Week_Shift'] = '0'
            df.loc[row_idx, 'WO_Sequence'] = current_week
            df.loc[row_idx, 'Days_Since_Last_WO'] = days_in_week
            df.loc[row_idx, 'Week_Number'] = current_week
            continue
        
        # Skip if Remarks is "Left"
        if pd.notna(remarks) and str(remarks).strip().lower() == 'left':
            df.loc[row_idx, 'Calculated_WO'] = 'N'
            df.loc[row_idx, 'Week_Shift'] = shift if pd.notna(shift) else ''
            df.loc[row_idx, 'WO_Sequence'] = current_week
            df.loc[row_idx, 'Days_Since_Last_WO'] = days_in_week
            df.loc[row_idx, 'Week_Number'] = current_week
            continue
        
        # Check if we need to assign WO (every 6 days)
        if days_in_week >= 6 or is_source_wo:
            # This is a WO day - end of current week
            
            # Assign the most common shift to the entire previous week
            if len(week_shifts) > 0 and len(week_indices) > 0:
                shift_counts = Counter(week_shifts)
                most_common_shift = shift_counts.most_common(1)[0][0]
                
                # Update all records in the week with the same shift
                for w_idx in week_indices:
                    if df.loc[w_idx, 'Calculated_WO'] != 'Y':  # Don't change WO records
                        df.loc[w_idx, 'Week_Shift'] = most_common_shift
            
            # Mark this as WO
            df.loc[row_idx, 'Calculated_WO'] = 'Y'
            df.loc[row_idx, 'Week_Shift'] = 'WO'
            df.loc[row_idx, 'WO_Sequence'] = current_week + 1
            df.loc[row_idx, 'Days_Since_Last_WO'] = 0
            df.loc[row_idx, 'Week_Number'] = current_week + 1
            
            # Reset for next week
            current_week += 1
            days_in_week = 0
            week_shifts = []
            week_indices = []
        else:
            # Regular day in the week
            df.loc[row_idx, 'Calculated_WO'] = 'N'
            df.loc[row_idx, 'WO_Sequence'] = current_week
            df.loc[row_idx, 'Days_Since_Last_WO'] = days_in_week
            df.loc[row_idx, 'Week_Number'] = current_week
            
            # Track shift for max calculation
            if pd.notna(shift) and str(shift) not in ['0', 'WO'] and str(shift) != 'nan':
                week_shifts.append(str(shift))
            week_indices.append(row_idx)
            
            days_in_week += 1
    
    # Handle the last week - assign most common shift to remaining records
    if len(week_shifts) > 0 and len(week_indices) > 0:
        shift_counts = Counter(week_shifts)
        most_common_shift = shift_counts.most_common(1)[0][0]
        
        for w_idx in week_indices:
            if df.loc[w_idx, 'Calculated_WO'] != 'Y':
                df.loc[w_idx, 'Week_Shift'] = most_common_shift

print("\nProcessing complete!")

# Now update First In, Last Out, Hrs based on Week_Shift from Sheet2
print("\nUpdating shift timings from Sheet2...")

for idx, row in df.iterrows():
    week_shift = row['Week_Shift']
    if week_shift and week_shift in shift_timings and week_shift != 'WO':
        timings = shift_timings[week_shift]
        df.loc[idx, 'First In'] = timings['First In']
        df.loc[idx, 'Last Out'] = timings['Last Out']
        df.loc[idx, 'Hrs'] = timings['Hrs']

# Display summary
print("\n=== SUMMARY ===")
print(f"Total records: {len(df)}")
print(f"WO records assigned: {(df['Calculated_WO'] == 'Y').sum()}")
print(f"Unique WO sequences: {df['WO_Sequence'].nunique()}")

# Show shift distribution in Week_Shift
print(f"\nWeek_Shift distribution:\n{df['Week_Shift'].value_counts()}")

# Show sample of results for first employee
print("\n=== SAMPLE OUTPUT (First 25 records of first employee) ===")
sample_cols = ['Empl Id', 'Attendance Date', 'Shift', 'Remarks', 'Duty Status', 
               'Calculated_WO', 'Week_Shift', 'First In', 'Last Out', 'Hrs', 
               'WO_Sequence', 'Days_Since_Last_WO', 'Week_Number']

first_emp = df['Empl Id'].iloc[0]
print(df[df['Empl Id'] == first_emp][sample_cols].head(25).to_string())

# Show sample where WO is assigned
wo_assigned = df[df['Calculated_WO'] == 'Y']
if len(wo_assigned) > 0:
    print(f"\n=== SAMPLE WO ASSIGNMENTS (First 10) ===")
    print(wo_assigned[sample_cols].head(10).to_string())

# Save to new Excel file
print(f"\nSaving to: {output_path}")
df.to_excel(output_path, index=False, sheet_name='Processed_Data')
print("Done!")
