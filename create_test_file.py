#!/usr/bin/env python3
"""Create a test file with 3 employees for testing the WO allocation logic."""

import pandas as pd
from datetime import datetime, timedelta

# Create sample data for 3 employees across multiple weeks
# Each employee has different shift patterns

data = []

# Employee 1: Mix of AS, CS, GS shifts with some 0s and leaves
emp1_id = 11104
start_date = datetime(2025, 5, 24)
emp1_shifts = ['AS', '0', 'CS', 'CS', 'GS', 'AS', 'WO',  # Week 1 ends with WO
               '0', 'BS', '0', 'CL', 'CL', 'BS', 'WO',   # Week 2 (with leaves)
               'AS', 'AS', 'AS', 'AS', 'AS', 'AS', 'WO', # Week 3 (all AS)
               'CS', 'CS', 'CS', 'CS', 'CS', 'CS', 'WO'] # Week 4 (all CS)

for i, shift in enumerate(emp1_shifts):
    date = start_date + timedelta(days=i)
    duty_status = 'WO' if shift == 'WO' else 'Y'
    remarks = None
    
    # Set remarks for CL (Casual Leave)
    if shift == 'CL':
        remarks = 'Casual Leave'
        duty_status = 'CL'
    elif shift == '0':
        duty_status = 'WO' if i == 6 else '0'
    
    data.append({
        'Month': 'June',
        'Empl Id': emp1_id,
        'Attendance Date': date.strftime('%Y-%m-%d'),
        'Duty Status': duty_status,
        'Shift': shift if shift not in ['WO', 'CL'] else ('0' if shift == '0' else None),
        'First In': '08:00:00' if shift not in ['WO', '0', 'CL'] else ('00:00:00' if shift in ['0', 'WO'] else None),
        'Last Out': '17:00:00' if shift not in ['WO', '0', 'CL'] else ('00:00:00' if shift in ['0', 'WO'] else None),
        'Hrs': '08:00:00' if shift not in ['WO', '0', 'CL'] else ('00:00:00' if shift in ['0', 'WO'] else None),
        'OT Hours': 0,
        'OT Status': 'A',
        'Half Day Type': None,
        'Remarks': remarks,
        'Check': None
    })

# Employee 2: Mostly BS shifts
emp2_id = 11445
emp2_shifts = ['BS', 'BS', 'BS', 'BS', 'BS', 'BS', 'WO',  # Week 1 (all BS)
               'GS', 'GS', 'GS', 'GS', 'GS', 'GS', 'WO',  # Week 2 (all GS)
               'AS', 'AS', '0', 'AS', 'AS', 'AS', 'WO',   # Week 3 (with one 0)
               'CS', 'CS', 'CS', 'CS', 'CS', 'CS', 'WO']  # Week 4 (all CS)

for i, shift in enumerate(emp2_shifts):
    date = start_date + timedelta(days=i)
    duty_status = 'WO' if shift == 'WO' else 'Y'
    
    data.append({
        'Month': 'June',
        'Empl Id': emp2_id,
        'Attendance Date': date.strftime('%Y-%m-%d'),
        'Duty Status': duty_status,
        'Shift': shift if shift != 'WO' else '0',
        'First In': '08:00:00' if shift not in ['WO', '0'] else '00:00:00',
        'Last Out': '17:00:00' if shift not in ['WO', '0'] else '00:00:00',
        'Hrs': '08:00:00' if shift not in ['WO', '0'] else '00:00:00',
        'OT Hours': 0,
        'OT Status': 'A',
        'Half Day Type': None,
        'Remarks': None,
        'Check': None
    })

# Employee 3: Mixed with "Left" remarks
emp3_id = 11447
emp3_shifts = ['GS', 'GS', 'GS', 'GS', 'GS', 'GS', 'WO',  # Week 1
               'AS', 'AS', 'Left', 'AS', 'AS', 'AS', 'WO', # Week 2 (one Left)
               'BS', 'BS', 'BS', 'BS', 'BS', 'BS', 'WO',    # Week 3
               'CS', '0', 'CS', 'CS', 'CS', 'CS', 'WO']     # Week 4 (one 0)

emp3_remarks = [None, None, None, None, None, None, None,
                None, None, 'Left', None, None, None, None,
                None, None, None, None, None, None, None,
                None, None, None, None, None, None, None]

for i, shift in enumerate(emp3_shifts):
    date = start_date + timedelta(days=i)
    duty_status = 'WO' if shift == 'WO' else 'Y'
    
    actual_shift = shift
    if shift == 'Left':
        actual_shift = 'AS'  # Original shift before leaving
        duty_status = 'N'
    elif shift == 'WO':
        actual_shift = '0'
    
    data.append({
        'Month': 'June',
        'Empl Id': emp3_id,
        'Attendance Date': date.strftime('%Y-%m-%d'),
        'Duty Status': duty_status,
        'Shift': actual_shift,
        'First In': '08:00:00' if shift not in ['WO', '0', 'Left'] else '00:00:00',
        'Last Out': '17:00:00' if shift not in ['WO', '0', 'Left'] else '00:00:00',
        'Hrs': '08:00:00' if shift not in ['WO', '0', 'Left'] else '00:00:00',
        'OT Hours': 0,
        'OT Status': 'A',
        'Half Day Type': None,
        'Remarks': emp3_remarks[i],
        'Check': None
    })

# Create DataFrame
df_test = pd.DataFrame(data)

# Save to Excel with two sheets (Sheet1 for data, Sheet2 for shift timings)
output_file = '/Users/smartmysql/Documents/sukumarHR/test_audit.xlsx'

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    df_test.to_excel(writer, sheet_name='Sheet1', index=False)
    
    # Create Sheet2 with shift timings (same structure as original)
    shift_data = {
        'GS': ['08:50', '17:30', '08:40'],
        'AS': ['05:54', '14:03', '08:09'],
        'CS': ['21:55', '06:03', '08:08'],
        'BS': ['13:57', '22:03', '08:06']
    }
    
    # Create Sheet2 with proper structure
    sheet2_rows = []
    for i in range(10):  # 10 rows like original
        row = {}
        col_idx = 0
        for shift_code, timings in shift_data.items():
            row[f'Col_{col_idx}'] = shift_code
            row[f'Col_{col_idx+1}'] = timings[0]
            row[f'Col_{col_idx+2}'] = timings[1]
            row[f'Col_{col_idx+3}'] = timings[2]
            row[f'Col_{col_idx+4}'] = None  # Empty separator
            col_idx += 5
        sheet2_rows.append(row)
    
    df_sheet2 = pd.DataFrame(sheet2_rows)
    df_sheet2.to_excel(writer, sheet_name='Sheet2', index=False, header=False)

print(f"Test file created: {output_file}")
print(f"Total records: {len(df_test)}")
print(f"Employees: {df_test['Empl Id'].unique()}")
print("\nSample data:")
print(df_test.head(20).to_string())
