#!/usr/bin/env python3
"""Check the 13-day gap between WOs."""

import pandas as pd

file_path = "/Users/smartmysql/Documents/sukumarHR/Audit Work March 50_7day.xlsx"
df = pd.read_excel(file_path, sheet_name='Processed_Data')

# Check employee 11104 for the specific 13-day gap
emp_id = 11104
emp_data = df[df['Empl Id'] == emp_id].copy()
emp_data = emp_data.sort_values('Attendance Date').reset_index(drop=True)

# Focus on rows around the 13-day gap (rows 85-110)
print("Employee 11104 - Rows 85 to 110:")
print("="*80)
cols_to_show = ['Row_Num', 'Attendance Date', 'Shift', 'Duty Status', 'Calculated_WO', 
               'Week_Shift', 'Duty Code', 'WO_Sequence', 'Days_Since_Last_WO']

# Add row numbers
emp_data['Row_Num'] = emp_data.index + 2  # Excel row numbers

subset = emp_data[(emp_data['Row_Num'] >= 85) & (emp_data['Row_Num'] <= 110)]
print(subset[cols_to_show].to_string())

print("\n" + "="*80)
print("Analysis of the gap:")
print("="*80)

# Find the specific gap
prev_wo = None
curr_wo = None

for idx, row in subset.iterrows():
    if row['Calculated_WO'] == 'Y':
        if prev_wo is None:
            prev_wo = row
        else:
            curr_wo = row
            break

if prev_wo and curr_wo:
    days_gap = curr_wo['Row_Num'] - prev_wo['Row_Num']
    print(f"Previous WO: Row {prev_wo['Row_Num']}, Date: {prev_wo['Attendance Date'].strftime('%Y-%m-%d')}")
    print(f"Current WO: Row {curr_wo['Row_Num']}, Date: {curr_wo['Attendance Date'].strftime('%Y-%m-%d')}")
    print(f"Days gap: {days_gap}")
    
    print(f"\nChecking for WOs marked as Absent between these dates:")
    absent_wo = subset[(subset['Duty Status'] == 'WO') & (subset['Duty Code'] == 'Absent')]
    if not absent_wo.empty:
        print("Found WOs marked as Absent:")
        print(absent_wo[['Row_Num', 'Attendance Date', 'Duty Code']].to_string())
    else:
        print("No WOs marked as Absent in this period")
    
    print(f"\nChecking Days_Since_Last_WO progression:")
    for idx, row in subset.iterrows():
        if row['Row_Num'] >= prev_wo['Row_Num'] and row['Row_Num'] <= curr_wo['Row_Num']:
            print(f"  Row {row['Row_Num']:3d}: Days_Since_Last_WO = {row['Days_Since_Last_WO']}")
