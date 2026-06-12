#!/usr/bin/env python3
"""Debug employee 11104 around row 99."""

import pandas as pd

file_path = "/Users/smartmysql/Documents/sukumarHR/Audit Work March 50_7day.xlsx"
df = pd.read_excel(file_path, sheet_name='Processed_Data')

# Filter for employee 11104
emp_11104 = df[df['Empl Id'] == 11104].copy()
emp_11104['Row_Num'] = emp_11104.index + 1  # Excel row number (1-based)

# Show rows around 99 (rows 90-110)
print("Employee 11104 - Rows 90 to 110:")
print("="*80)
cols_to_show = ['Row_Num', 'Attendance Date', 'Shift', 'Duty Status', 'Calculated_WO', 
               'Week_Shift', 'Duty Code', 'WO_Sequence', 'Days_Since_Last_WO']
subset = emp_11104[(emp_11104['Row_Num'] >= 90) & (emp_11104['Row_Num'] <= 110)]
print(subset[cols_to_show].to_string())

# Check WO positions and distances
print("\n" + "="*80)
print("WO Analysis for Employee 11104:")
print("="*80)

wo_rows = emp_11104[emp_11104['Duty Status'] == 'WO'][['Row_Num', 'Attendance Date', 'Calculated_WO', 'Duty Code']]
print("Original WO positions:")
print(wo_rows.to_string())

# Calculate distances between WOs
wo_positions = wo_rows['Row_Num'].tolist()
print(f"\nWO positions: {wo_positions}")
print("Distances between consecutive WOs:")
for i in range(len(wo_positions) - 1):
    distance = wo_positions[i+1] - wo_positions[i]
    print(f"  Row {wo_positions[i]} to Row {wo_positions[i+1]}: {distance} days")
    if distance <= 7:
        print(f"    -> Multiple WOs in 7-day window! Later WO marked as Absent")

# Check what's at row 99 specifically
print("\n" + "="*80)
print("Row 99 Details:")
print("="*80)
row_99 = emp_11104[emp_11104['Row_Num'] == 99]
if not row_99.empty:
    print(row_99[cols_to_show].to_string())
else:
    print("Row 99 not found for employee 11104")
