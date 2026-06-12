#!/usr/bin/env python3
"""Debug the 13-day gap issue."""

import pandas as pd

file_path = "/Users/smartmysql/Documents/sukumarHR/Audit Work March 50_7day.xlsx"
df = pd.read_excel(file_path, sheet_name='Processed_Data')

# Check employee 11104 for the 13-day gap
emp_id = 11104
emp_data = df[df['Empl Id'] == emp_id].copy()
emp_data = emp_data.sort_values('Attendance Date').reset_index(drop=True)

# Find WOs
wo_positions = []
for idx, row in emp_data.iterrows():
    if row['Calculated_WO'] == 'Y':
        wo_positions.append(idx)

print(f"Employee {emp_id} WO positions:")
for i, pos in enumerate(wo_positions):
    date = emp_data.loc[pos, 'Attendance Date']
    print(f"  WO {i+1}: Row {pos+1}, Date: {date.strftime('%Y-%m-%d')}")

# Find the 13-day gap
for i in range(1, len(wo_positions)):
    prev_wo = wo_positions[i-1]
    curr_wo = wo_positions[i]
    days_between = curr_wo - prev_wo
    
    if days_between == 13:
        print(f"\nFound 13-day gap between WO {i} and WO {i+1}:")
        print(f"  Previous WO: Row {prev_wo+1}, {emp_data.loc[prev_wo, 'Attendance Date'].strftime('%Y-%m-%d')}")
        print(f"  Current WO: Row {curr_wo+1}, {emp_data.loc[curr_wo, 'Attendance Date'].strftime('%Y-%m-%d')}")
        
        print(f"\nDays between these WOs:")
        for j in range(prev_wo + 1, curr_wo):
            date = emp_data.loc[j, 'Attendance Date']
            shift = emp_data.loc[j, 'Shift']
            week_shift = emp_data.loc[j, 'Week_Shift']
            duty_status = emp_data.loc[j, 'Duty Status']
            calculated_wo = emp_data.loc[j, 'Calculated_WO']
            duty_code = emp_data.loc[j, 'Duty Code']
            
            print(f"  Row {j+1:3d}: {date.strftime('%Y-%m-%d')} - Shift: {shift:2s} -> Week_Shift: {week_shift:2s}, "
                  f"Duty: {duty_status:7s} -> {calculated_wo}, Duty Code: {duty_code}")
        
        # Check if there was a WO marked as Absent in between
        print(f"\nChecking for WOs marked as Absent in this period:")
        for j in range(prev_wo + 1, curr_wo):
            if emp_data.loc[j, 'Duty Status'] == 'WO' and emp_data.loc[j, 'Duty Code'] == 'Absent':
                date = emp_data.loc[j, 'Attendance Date']
                print(f"  Found WO marked as Absent at Row {j+1}: {date.strftime('%Y-%m-%d')}")
        
        break
