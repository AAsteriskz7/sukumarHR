#!/usr/bin/env python3
"""Generate WO every 7 days. Multiple WOs in 7-day window = Absent."""

import pandas as pd
from collections import Counter
import random

file_path = "/Users/smartmysql/Documents/sukumarHR/Audit Work March 50.xlsx"
output_path = "/Users/smartmysql/Documents/sukumarHR/Audit Work March 50_7day.xlsx"

print(f"Reading: {file_path}")
df = pd.read_excel(file_path, sheet_name='23', header=0)
df_shifts = pd.read_excel(file_path, sheet_name='Sheet2', header=None)

# Extract ALL shift timing rows from Sheet2
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
df['Duty Code'] = ''  # New column for duty code

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
    
    # Check for multiple WOs in any 7-day window
    wo_positions = []
    for i in range(len(emp_indices)):
        row_idx = emp_indices[i]
        duty_status = df.loc[row_idx, 'Duty Status']
        if pd.notna(duty_status) and str(duty_status).strip().upper() == 'WO':
            wo_positions.append(i)
    
    # Mark extra WOs as Absent
    marked_absent = []
    for i, wo_pos in enumerate(wo_positions):
        # Check if there's another WO within next 6 positions (less than 7 days)
        for j in range(i + 1, len(wo_positions)):
            if wo_positions[j] - wo_pos < 7:
                # Mark the later WO as Absent
                marked_absent.append(wo_positions[j])
    
    print(f"  Found {len(wo_positions)} WOs, marking {len(marked_absent)} as Absent")
    
    # Mark first WO
    if first_wo_idx < len(emp_indices):
        df.loc[emp_indices[first_wo_idx], 'Calculated_WO'] = 'Y'
        df.loc[emp_indices[first_wo_idx], 'Week_Shift'] = 'WO'
        df.loc[emp_indices[first_wo_idx], 'WO_Sequence'] = 1
        df.loc[emp_indices[first_wo_idx], 'Days_Since_Last_WO'] = 0
        df.loc[emp_indices[first_wo_idx], 'Duty Code'] = 'WO'
    
    # Process: count exactly 6 days, then assign WO on 7th day
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
        is_marked_absent = i in marked_absent
        
        days_counter += 1
        
        # Check if this should be WO (after exactly 6 working days)
        if days_in_week >= 6 or (is_source_wo and not is_marked_absent):
            # Assign most common shift to week
            if week_shifts and week_indices:
                most_common = Counter(week_shifts).most_common(1)[0][0]
                for w_idx in week_indices:
                    if not should_skip(df.loc[w_idx, 'Shift'], df.loc[w_idx, 'Remarks']):
                        df.loc[w_idx, 'Week_Shift'] = most_common
                    else:
                        orig_shift = df.loc[w_idx, 'Shift']
                        df.loc[w_idx, 'Week_Shift'] = str(orig_shift) if pd.notna(orig_shift) else ''
            
            # Mark this as WO or Absent
            if is_marked_absent:
                df.loc[row_idx, 'Calculated_WO'] = 'N'
                df.loc[row_idx, 'Week_Shift'] = '0'
                df.loc[row_idx, 'WO_Sequence'] = current_week
                df.loc[row_idx, 'Days_Since_Last_WO'] = days_counter
                df.loc[row_idx, 'Duty Code'] = 'Absent'
                print(f"  Day {i}: Marked as Absent (multiple WOs in 7-day window)")
                # DON'T increment days_in_week for Absent WOs - they don't count towards the 6-day cycle
                # DON'T reset - continue counting as regular day for days_counter
            else:
                df.loc[row_idx, 'Calculated_WO'] = 'Y'
                df.loc[row_idx, 'Week_Shift'] = 'WO'
                df.loc[row_idx, 'WO_Sequence'] = current_week + 1
                df.loc[row_idx, 'Days_Since_Last_WO'] = 0
                df.loc[row_idx, 'Duty Code'] = 'WO'
                print(f"  Day {i}: WO assigned (days_in_week={days_in_week})")
                
                # Reset only for valid WO
                current_week += 1
                days_in_week = 0
                week_shifts = []
                week_indices = []
                days_counter = 0
        else:
            # Regular day
            if is_marked_absent:
                df.loc[row_idx, 'Calculated_WO'] = 'N'
                df.loc[row_idx, 'Week_Shift'] = '0'
                df.loc[row_idx, 'WO_Sequence'] = current_week
                df.loc[row_idx, 'Days_Since_Last_WO'] = days_counter
                df.loc[row_idx, 'Duty Code'] = 'Absent'
            else:
                df.loc[row_idx, 'Calculated_WO'] = 'N'
                df.loc[row_idx, 'WO_Sequence'] = current_week
                df.loc[row_idx, 'Days_Since_Last_WO'] = days_counter
                df.loc[row_idx, 'Duty Code'] = df.loc[row_idx, 'Duty Status']  # Keep original duty status
            
            # Track for shift calculation (only for non-Absent days)
            if not is_marked_absent:
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

# Add new columns
df['new_First In'] = df['First In']
df['new_Last Out'] = df['Last Out']
df['new_Hrs'] = df['Hrs']

# Update new columns based on shift comparison
print("\nUpdating new_First In, new_Last Out, new_Hrs columns...")
updated_count = 0
preserved_count = 0

for idx, row in df.iterrows():
    original_shift = str(row['Shift']) if pd.notna(row['Shift']) else ''
    week_shift = str(row['Week_Shift']) if pd.notna(row['Week_Shift']) else ''
    
    # Skip if should preserve
    if should_skip(row['Shift'], row['Remarks']):
        preserved_count += 1
        continue
    
    # Skip if Week_Shift is WO, 0, or empty
    if not week_shift or week_shift in ['WO', '0', '']:
        preserved_count += 1
        continue
    
    # Check if shift changed
    if original_shift and original_shift != '0' and original_shift != week_shift:
        # Shift changed - get random row from Sheet2 for new shift
        matching_rows = [r for r in shift_timing_rows if r['shift'] == week_shift]
        if matching_rows:
            random_row = random.choice(matching_rows)
            df.loc[idx, 'new_First In'] = random_row['first_in']
            df.loc[idx, 'new_Last Out'] = random_row['last_out']
            df.loc[idx, 'new_Hrs'] = random_row['hrs']
            updated_count += 1
    else:
        # Shift didn't change - keep original values in new columns
        preserved_count += 1

print(f"Updated {updated_count} records with new shift timings")
print(f"Preserved {preserved_count} records with original values")

print(f"\nDone! WO assigned: {(df['Calculated_WO'] == 'Y').sum()}")
print(f"Absent marked: {(df['Duty Code'] == 'Absent').sum()}")
print(f"Saving to: {output_path}")
df.to_excel(output_path, index=False, sheet_name='Processed_Data')
print("Saved!")
