#!/usr/bin/env python3
"""Check the 7-day and 5-day week issues."""

import pandas as pd

file_path = "/Users/smartmysql/Documents/sukumarHR/Audit Work March 50_7day.xlsx"
df = pd.read_excel(file_path, sheet_name='Processed_Data')

# Check employee 11104 around the issue area
emp_id = 11104
emp_data = df[df['Empl Id'] == emp_id].copy()
emp_data = emp_data.sort_values('Attendance Date').reset_index(drop=True)
emp_data['Row_Num'] = emp_data.index + 2

# Find the area with 7-day week
print("Employee 11104 - Looking for 7-day week issue:")
print("="*80)

# Find WOs and check the gaps
wo_positions = []
for idx, row in emp_data.iterrows():
    if row['Calculated_WO'] == 'Y':
        wo_positions.append(idx)

for i in range(1, len(wo_positions)):
    prev_wo = wo_positions[i-1]
    curr_wo = wo_positions[i]
    days_between = curr_wo - prev_wo
    
    if days_between == 8:  # 7 days between WOs
        print(f"\nFound 7-day gap between WO {i} and WO {i+1}:")
        print(f"  Previous WO: Row {prev_wo+1}, {emp_data.loc[prev_wo, 'Attendance Date'].strftime('%Y-%m-%d')}")
        print(f"  Current WO: Row {curr_wo+1}, {emp_data.loc[curr_wo, 'Attendance Date'].strftime('%Y-%m-%d')}")
        
        print(f"\nDays between these WOs:")
        for j in range(prev_wo + 1, curr_wo):
            date = emp_data.loc[j, 'Attendance Date']
            shift = emp_data.loc[j, 'Week_Shift']
            duty_code = emp_data.loc[j, 'Duty Code']
            days_since = emp_data.loc[j, 'Days_Since_Last_WO']
            
            print(f"  Row {j+1:3d}: {date.strftime('%Y-%m-%d')} - Week_Shift: {shift:2s}, "
                  f"Duty Code: {duty_code:7s}, Days_Since: {days_since}")
        
        # Check if there was an Absent WO
        for j in range(prev_wo + 1, curr_wo):
            if emp_data.loc[j, 'Duty Code'] == 'Absent':
                print(f"\n  -> Found Absent at Row {j+1}, which counts as a day but doesn't reset counter")
        break

# Now check for 5-day week
print("\n" + "="*80)
print("Looking for 5-day week issue:")
print("="*80)

for i in range(1, len(wo_positions)):
    prev_wo = wo_positions[i-1]
    curr_wo = wo_positions[i]
    days_between = curr_wo - prev_wo
    
    if days_between == 6:  # 5 days between WOs
        print(f"\nFound 5-day gap between WO {i} and WO {i+1}:")
        print(f"  Previous WO: Row {prev_wo+1}, {emp_data.loc[prev_wo, 'Attendance Date'].strftime('%Y-%m-%d')}")
        print(f"  Current WO: Row {curr_wo+1}, {emp_data.loc[curr_wo, 'Attendance Date'].strftime('%Y-%m-%d')}")
        
        print(f"\nDays between these WOs:")
        for j in range(prev_wo + 1, curr_wo):
            date = emp_data.loc[j, 'Attendance Date']
            shift = emp_data.loc[j, 'Week_Shift']
            duty_code = emp_data.loc[j, 'Duty Code']
            days_since = emp_data.loc[j, 'Days_Since_Last_WO']
            
            print(f"  Row {j+1:3d}: {date.strftime('%Y-%m-%d')} - Week_Shift: {shift:2s}, "
                  f"Duty Code: {duty_code:7s}, Days_Since: {days_since}")
        break
