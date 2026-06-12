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

# Find WOs marked as Absent in this period
print("WOs marked as Absent in this period:")
absent_wo = subset[(subset['Duty Status'] == 'WO') & (subset['Duty Code'] == 'Absent')]
if not absent_wo.empty:
    print(absent_wo[['Row_Num', 'Attendance Date', 'Duty Code']].to_string())
    
    print("\nProblem identified:")
    print("1. Row 93: Original WO marked as Absent (Days_Since_Last_WO = 3)")
    print("2. Row 97: Original WO marked as Absent (Days_Since_Last_WO = 7)")
    print("3. These WOs were too close to previous WO (less than 7 days)")
    print("4. After Row 97, counter reset and started counting from 1 again")
    print("5. Next WO appeared at Row 104 after 6 more days")
    
    print("\nExpected behavior:")
    print("- When WOs are marked as Absent, they should still count towards the 6-day cycle")
    print("- The script should continue counting days even when marking WOs as Absent")
else:
    print("No WOs marked as Absent in this period")
