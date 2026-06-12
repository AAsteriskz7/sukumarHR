#!/usr/bin/env python3
"""Test random shift selection on test_50.xlsx."""

import pandas as pd
from collections import Counter
import random

file_path = "/Users/smartmysql/Documents/sukumarHR/test_50.xlsx"
output_path = "/Users/smartmysql/Documents/sukumarHR/test_50_v3_processed.xlsx"

print(f"Reading test file: {file_path}")
df = pd.read_excel(file_path, sheet_name='23', header=0)
df_shifts = pd.read_excel(file_path, sheet_name='Sheet2', header=None)

# Extract ALL shift timing rows
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

print(f"Shift timing rows collected: {len(shift_timing_rows)}")
for shift in ['GS', 'AS', 'CS', 'BS']:
    count = sum(1 for r in shift_timing_rows if r['shift'] == shift)
    print(f"  {shift}: {count} rows")

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
    
    print(f"\nProcessing Employee {emp_id}")
    
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
    
    # Process
    current_week = 1
    days_in_week = 0
    week_shifts = []
    week_indices = []
    days_counter = 0
    
    for i in range(first_wo_idx + 1, len(emp_indices)):
        row_idx = emp_indices[i]
        shift = df.loc[row_idx, 'Shift']
        remarks = df.loc[row_idx, 'Remarks']
        duty_status = df.loc[row_idx, 'Duty Status']
        
        is_source_wo = pd.notna(duty_status) and str(duty_status).strip().upper() == 'WO'
        
        days_counter += 1
        
        if days_in_week >= 6 or is_source_wo:
            # Assign most common shift
            if week_shifts and week_indices:
                most_common = Counter(week_shifts).most_common(1)[0][0]
                for w_idx in week_indices:
                    if not should_skip(df.loc[w_idx, 'Shift'], df.loc[w_idx, 'Remarks']):
                        df.loc[w_idx, 'Week_Shift'] = most_common
                    else:
                        orig_shift = df.loc[w_idx, 'Shift']
                        df.loc[w_idx, 'Week_Shift'] = str(orig_shift) if pd.notna(orig_shift) else ''
            
            # Mark WO
            df.loc[row_idx, 'Calculated_WO'] = 'Y'
            df.loc[row_idx, 'Week_Shift'] = 'WO'
            df.loc[row_idx, 'WO_Sequence'] = current_week + 1
            df.loc[row_idx, 'Days_Since_Last_WO'] = 0
            
            # Reset
            current_week += 1
            days_in_week = 0
            week_shifts = []
            week_indices = []
            days_counter = 0
        else:
            # Regular day
            df.loc[row_idx, 'Calculated_WO'] = 'N'
            df.loc[row_idx, 'WO_Sequence'] = current_week
            df.loc[row_idx, 'Days_Since_Last_WO'] = days_counter
            
            if not should_skip(shift, remarks) and pd.notna(shift) and str(shift) not in ['0', 'WO', 'nan']:
                week_shifts.append(str(shift))
            week_indices.append(row_idx)
            days_in_week += 1
    
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
        else:
            for w_idx in week_indices:
                orig_shift = df.loc[w_idx, 'Shift']
                df.loc[w_idx, 'Week_Shift'] = str(orig_shift) if pd.notna(orig_shift) else ''

# Update First In/Last Out based on shift changes
print("\nUpdating First In/Last Out columns...")
updated_count = 0
preserved_count = 0

for idx, row in df.iterrows():
    original_shift = row['Shift']
    week_shift = row['Week_Shift']
    
    if should_skip(original_shift, row['Remarks']):
        preserved_count += 1
        continue
    
    if not week_shift or week_shift in ['WO', '0', '']:
        preserved_count += 1
        continue
    
    if pd.notna(original_shift) and str(original_shift) != '0' and str(original_shift) != week_shift:
        # Shift changed - get random timing
        matching_rows = [r for r in shift_timing_rows if r['shift'] == week_shift]
        if matching_rows:
            random_row = random.choice(matching_rows)
            df.loc[idx, 'First In'] = random_row['first_in']
            df.loc[idx, 'Last Out'] = random_row['last_out']
            df.loc[idx, 'Hrs'] = random_row['hrs']
            updated_count += 1
            print(f"  Updated: {row['Empl Id']} {row['Attendance Date'].strftime('%Y-%m-%d')} {original_shift}->{week_shift}")
    else:
        preserved_count += 1

print(f"\nUpdated {updated_count} records with new shift timings")
print(f"Preserved {preserved_count} records with original values")

# Show results
print("\n" + "="*80)
print("DETAILED RESULTS (First 30 records)")
print("="*80)

sample_cols = ['Empl Id', 'Attendance Date', 'Shift', 'Week_Shift', 'First In', 'Last Out', 'Hrs', 
               'Calculated_WO', 'WO_Sequence', 'Days_Since_Last_WO']

print(df[sample_cols].head(30).to_string())

print(f"\nSaving to: {output_path}")
df.to_excel(output_path, index=False, sheet_name='Processed_Data')
print("Saved!")
