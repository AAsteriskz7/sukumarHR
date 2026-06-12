#!/usr/bin/env python3
"""Simple 7-day WO logic: 6 days (any type) + WO."""

import pandas as pd
from collections import Counter
import random

file_path = "/Users/smartmysql/Documents/sukumarHR/Audit Work March 50.xlsx"
output_path = "/Users/smartmysql/Documents/sukumarHR/Audit Work March 50_7day_simple.xlsx"

print(f"Reading: {file_path}")
df = pd.read_excel(file_path, sheet_name='23', header=0)
df_shifts = pd.read_excel(file_path, sheet_name='Sheet2', header=None)

# Extract shift timing rows
shift_timing_rows = []
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
                            hrs_val = int(hrs_str.split(':')[0]) + int(hrs_str.split(':')[1])/60 if ':' in hrs_str else 8.5
                            if hrs_val > 8:
                                shift_timing_rows.append({
                                    'shift': shift_code,
                                    'first_in': first_in,
                                    'last_out': last_out,
                                    'hrs': hrs
                                })
                        except:
                            pass
        col_idx += 1

print(f"Shift timing rows: {len(shift_timing_rows)}")

# Convert and sort
df['Attendance Date'] = pd.to_datetime(df['Attendance Date'])
df = df.sort_values(['Empl Id', 'Attendance Date']).reset_index(drop=True)

# New columns
df['Calculated_WO'] = 'N'
df['Week_Shift'] = ''
df['WO_Sequence'] = 0
df['Days_Since_Last_WO'] = 0

def should_skip(shift, remarks):
    if pd.notna(shift) and str(shift) == '0':
        return True
    if pd.notna(remarks) and str(remarks).strip().lower() == 'left':
        return True
    return False

# Process each employee
for emp_id in df['Empl Id'].unique():
    emp_indices = df[df['Empl Id'] == emp_id].index.tolist()
    if not emp_indices:
        continue
    
    # Find first WO
    first_wo_idx = None
    for i, row_idx in enumerate(emp_indices):
        if pd.notna(df.loc[row_idx, 'Duty Status']) and str(df.loc[row_idx, 'Duty Status']).strip().upper() == 'WO':
            first_wo_idx = i
            break
    
    if first_wo_idx is None:
        first_wo_idx = 0
    
    # Mark first WO
    if first_wo_idx < len(emp_indices):
        df.loc[emp_indices[first_wo_idx], 'Calculated_WO'] = 'Y'
        df.loc[emp_indices[first_wo_idx], 'Week_Shift'] = 'WO'
        df.loc[emp_indices[first_wo_idx], 'WO_Sequence'] = 1
        df.loc[emp_indices[first_wo_idx], 'Days_Since_Last_WO'] = 0
    
    # Simple logic: count 6 days, then WO on 7th day
    current_week = 1
    days_counter = 0
    week_shifts = []
    week_indices = []
    
    for i in range(first_wo_idx + 1, len(emp_indices)):
        row_idx = emp_indices[i]
        shift = df.loc[row_idx, 'Shift']
        remarks = df.loc[row_idx, 'Remarks']
        
        days_counter += 1
        
        # After 6 days, assign WO
        if days_counter >= 7:
            # Assign most common shift to previous 6 days
            if week_shifts and week_indices:
                most_common = Counter(week_shifts).most_common(1)[0][0]
                for w_idx in week_indices:
                    if not should_skip(df.loc[w_idx, 'Shift'], df.loc[w_idx, 'Remarks']):
                        df.loc[w_idx, 'Week_Shift'] = most_common
                    else:
                        orig_shift = df.loc[w_idx, 'Shift']
                        df.loc[w_idx, 'Week_Shift'] = str(orig_shift) if pd.notna(orig_shift) else ''
            
            # Mark this as WO
            df.loc[row_idx, 'Calculated_WO'] = 'Y'
            df.loc[row_idx, 'Week_Shift'] = 'WO'
            df.loc[row_idx, 'WO_Sequence'] = current_week + 1
            df.loc[row_idx, 'Days_Since_Last_WO'] = 0
            
            # Reset
            current_week += 1
            days_counter = 0
            week_shifts = []
            week_indices = []
        else:
            # Regular day (1-6)
            df.loc[row_idx, 'Calculated_WO'] = 'N'
            df.loc[row_idx, 'WO_Sequence'] = current_week
            df.loc[row_idx, 'Days_Since_Last_WO'] = days_counter
            
            # Track for shift calculation
            if not should_skip(shift, remarks) and pd.notna(shift) and str(shift) not in ['0', 'WO', 'nan']:
                week_shifts.append(str(shift))
            week_indices.append(row_idx)
    
    # Final week
    if week_indices:
        if week_shifts:
            most_common = Counter(week_shifts).most_common(1)[0][0]
            for w_idx in week_indices:
                if not should_skip(df.loc[w_idx, 'Shift'], df.loc[w_idx, 'Remarks']):
                    df.loc[w_idx, 'Week_Shift'] = most_common
                else:
                    orig_shift = df.loc[w_idx, 'Shift']
                    df.loc[w_idx, 'Week_Shift'] = str(orig_shift) if pd.notna(orig_shift) else ''

# Add new columns
df['new_First In'] = df['First In']
df['new_Last Out'] = df['Last Out']
df['new_Hrs'] = df['Hrs']

# Update new columns
print("\nUpdating new_First In, new_Last Out, new_Hrs...")
for idx, row in df.iterrows():
    original_shift = str(row['Shift']) if pd.notna(row['Shift']) else ''
    week_shift = str(row['Week_Shift']) if pd.notna(row['Week_Shift']) else ''
    
    if should_skip(row['Shift'], row['Remarks']) or not week_shift or week_shift in ['WO', '0', '']:
        continue
    
    if original_shift and original_shift != '0' and original_shift != week_shift:
        matching_rows = [r for r in shift_timing_rows if r['shift'] == week_shift]
        if matching_rows:
            random_row = random.choice(matching_rows)
            df.loc[idx, 'new_First In'] = random_row['first_in']
            df.loc[idx, 'new_Last Out'] = random_row['last_out']
            df.loc[idx, 'new_Hrs'] = random_row['hrs']

print(f"\nWO assigned: {(df['Calculated_WO'] == 'Y').sum()}")
print(f"Saving to: {output_path}")
df.to_excel(output_path, index=False, sheet_name='Processed_Data')
print("Saved!")
