#!/usr/bin/env python3
"""Create test file from Audit Work March 50.xlsx with 3 employees."""

import pandas as pd
from datetime import datetime, timedelta

file_path = "/Users/smartmysql/Documents/sukumarHR/Audit Work March 50.xlsx"
output_file = "/Users/smartmysql/Documents/sukumarHR/test_50.xlsx"

print(f"Reading source file: {file_path}")

# Read the employee data sheet (named "23")
df_source = pd.read_excel(file_path, sheet_name='23', header=0)
print(f"Source shape: {df_source.shape}")
print(f"Columns: {df_source.columns.tolist()}")

# Get first 3 employees
unique_employees = df_source['Empl Id'].unique()[:3]
print(f"\nSelected employees: {unique_employees}")

# Filter data for these 3 employees
df_test = df_source[df_source['Empl Id'].isin(unique_employees)].copy()
print(f"Test data shape: {df_test.shape}")

# Sort by employee and date
df_test['Attendance Date'] = pd.to_datetime(df_test['Attendance Date'])
df_test = df_test.sort_values(['Empl Id', 'Attendance Date']).reset_index(drop=True)

print("\nTest data sample:")
print(df_test.head(20).to_string())

# Read Sheet2 for shift timings
df_shifts = pd.read_excel(file_path, sheet_name='Sheet2', header=None)
print(f"\nSheet2 shape: {df_shifts.shape}")

# Save to new Excel file
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    df_test.to_excel(writer, sheet_name='23', index=False)
    df_shifts.to_excel(writer, sheet_name='Sheet2', index=False, header=False)

print(f"\nTest file created: {output_file}")
print(f"Total records: {len(df_test)}")
print(f"Date range: {df_test['Attendance Date'].min()} to {df_test['Attendance Date'].max()}")

# Show employee summary
for emp in unique_employees:
    emp_data = df_test[df_test['Empl Id'] == emp]
    print(f"\nEmployee {emp}: {len(emp_data)} records")
    print(f"  Shifts: {emp_data['Shift'].unique()}")
    print(f"  Duty Status: {emp_data['Duty Status'].unique()}")
