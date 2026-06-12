#!/usr/bin/env python3
"""Test the corrected logic on test_50.xlsx - count all days for WO, preserve 0/Left for shifts."""

import pandas as pd
from collections import Counter

file_path = "/Users/smartmysql/Documents/sukumarHR/test_50.xlsx"
output_path = "/Users/smartmysql/Documents/sukumarHR/test_50_v2_processed.xlsx"

print(f"Reading test file: {file_path}")
df = pd.read_excel(file_path, sheet_name='23', header=0)
df_shifts = pd.read_excel(file_path, sheet_name='Sheet2', header=None)

# Extract shift timings
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
                            hrs_val = int(hrs_str.split(':')[0]) + int(hrs_str.split(':')[1])/60 if ':' in hrs_str else 8.5
                            if hrs_val > 8 and shift_code not in shift_timings:
                                shift_timings[shift_code] = {'First In': first_in, 'Last Out': last_out, 'Hrs': hrs}
                        except:
                            pass
        col_idx += 1

print(f"Shift timings: {shift_timings}")
print(f"Records: {len(df)}, Employees: {df['Empl Id'].nunique()}")

# Convert and sort
df['Attendance Date'] = pd.to_datetime(df['Attendance Date'])
df = df.sort_values(['Empl Id', 'Attendance Date']).reset_index(drop=True)

# New columns
df['Calculated_WO'] = 'N'
df['Week_Shift'] = ''
df['WO_Sequence'] = 0
df['Days_Since_Last_WO'] = 0

# Helper to check if record should preserve original
def preserve_original(shift, remarks):
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
    
    print(f"\nProcessing Employee {emp_id} ({len(emp_indices)} records)")
    
    # Find first WO
    first_wo_idx = None
    for i, row_idx in enumerate(emp_indices):
        if pd.notna(df.loc[row_idx, 'Duty Status']) and str(df.loc[row_idx, 'Duty Status']).strip().upper() == 'WO':
            first_wo_idx = i
            print(f"  First WO at position {i}")
            break
    
    if first_wo_idx is None:
        first_wo_idx = 0
        print(f"  No WO found, starting from position 0")
    
    # Mark first WO
    if first_wo_idx < len(emp_indices):
        df.loc[emp_indices[first_wo_idx], 'Calculated_WO'] = 'Y'
        df.loc[emp_indices[first_wo_idx], 'Week_Shift'] = 'WO'
        df.loc[emp_indices[first_wo_idx], 'WO_Sequence'] = 1
        df.loc[emp_indices[first_wo_idx], 'Days_Since_Last_WO'] = 0
    
    # Process: count exactly 6 days (including 0/Left), then assign WO
    current_week = 1
    days_in_week = 0  # Count ALL days
    week_shifts = []  # Only non-preserved for shift calculation
    week_indices = []
    days_counter = 0
    
    for i in range(first_wo_idx + 1, len(emp_indices)):
        row_idx = emp_indices[i]
        shift = df.loc[row_idx, 'Shift']
        remarks = df.loc[row_idx, 'Remarks']
        duty_status = df.loc[row_idx, 'Duty Status']
        date = df.loc[row_idx, 'Attendance Date']
        
        is_source_wo = pd.notna(duty_status) and str(duty_status).strip().upper() == 'WO'
        should_preserve = preserve_original(shift, remarks)
        
        days_counter += 1
        
        # Check if this should be WO (after exactly 6 days)
        if days_in_week >= 6 or is_source_wo:
            print(f"  Day {i} ({date.strftime('%Y-%m-%d')}): WO assigned (days_in_week={days_in_week})")
            
            # Assign most common shift to week (only from non-preserved)
            if week_shifts and week_indices:
                most_common = Counter(week_shifts).most_common(1)[0][0]
                print(f"    Week {current_week}: {dict(Counter(week_shifts))} -> '{most_common}'")
                for w_idx in week_indices:
                    if not preserve_original(df.loc[w_idx, 'Shift'], df.loc[w_idx, 'Remarks']):
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
            days_in_week = 0
            week_shifts = []
            week_indices = []
            days_counter = 0
        else:
            # Regular day
            df.loc[row_idx, 'Calculated_WO'] = 'N'
            df.loc[row_idx, 'WO_Sequence'] = current_week
            df.loc[row_idx, 'Days_Since_Last_WO'] = days_counter
            
            # Track for shift calculation
            if not should_preserve and pd.notna(shift) and str(shift) not in ['0', 'WO', 'nan']:
                week_shifts.append(str(shift))
            week_indices.append(row_idx)
            days_in_week += 1
            
            status = f"SKIP ({'Shift=0' if str(shift)=='0' else 'Remarks=Left'})" if should_preserve else f"Shift {shift}"
            print(f"  Day {i} ({date.strftime('%Y-%m-%d')}): {status} (days_in_week={days_in_week})")
    
    # Final week
    if week_indices:
        if week_shifts:
            most_common = Counter(week_shifts).most_common(1)[0][0]
            print(f"  Final week: {dict(Counter(week_shifts))} -> '{most_common}'")
            for w_idx in week_indices:
                if not preserve_original(df.loc[w_idx, 'Shift'], df.loc[w_idx, 'Remarks']):
                    df.loc[w_idx, 'Week_Shift'] = most_common
                else:
                    orig_shift = df.loc[w_idx, 'Shift']
                    df.loc[w_idx, 'Week_Shift'] = str(orig_shift) if pd.notna(orig_shift) else ''
        else:
            for w_idx in week_indices:
                orig_shift = df.loc[w_idx, 'Shift']
                df.loc[w_idx, 'Week_Shift'] = str(orig_shift) if pd.notna(orig_shift) else ''

# Apply shift timings
for idx, row in df.iterrows():
    ws = row['Week_Shift']
    if ws and ws in shift_timings and ws not in ['WO', '0']:
        if not preserve_original(row['Shift'], row['Remarks']):
            t = shift_timings[ws]
            df.loc[idx, 'First In'] = t['First In']
            df.loc[idx, 'Last Out'] = t['Last Out']
            df.loc[idx, 'Hrs'] = t['Hrs']

print(f"\nDone! WO assigned: {(df['Calculated_WO'] == 'Y').sum()}")
print(f"Week_Shift distribution:\n{df['Week_Shift'].value_counts()}")

# Show results
print("\n" + "="*80)
print("DETAILED RESULTS")
print("="*80)

sample_cols = ['Empl Id', 'Attendance Date', 'Shift', 'Remarks', 'Duty Status', 
               'Calculated_WO', 'Week_Shift', 'First In', 'Last Out', 'Hrs', 
               'WO_Sequence', 'Days_Since_Last_WO']

for emp_id in df['Empl Id'].unique():
    print(f"\n{'='*60}")
    print(f"EMPLOYEE {emp_id}")
    print('='*60)
    emp_data = df[df['Empl Id'] == emp_id][sample_cols]
    print(emp_data.to_string())

print(f"\nSaving to: {output_path}")
df.to_excel(output_path, index=False, sheet_name='Processed_Data')
print("Saved!")
