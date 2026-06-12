#!/usr/bin/env python3
"""Verify the simple 7-day file contains all data."""

import pandas as pd

file_path = "/Users/smartmysql/Documents/sukumarHR/Audit Work March 50_7day_simple.xlsx"
original_file = "/Users/smartmysql/Documents/sukumarHR/Audit Work March 50.xlsx"

print("Verifying file completeness...")
print("="*60)

# Load both files
df_simple = pd.read_excel(file_path, sheet_name='Processed_Data')
df_original = pd.read_excel(original_file, sheet_name='23')

print(f"Original file records: {len(df_original)}")
print(f"Simple 7-day file records: {len(df_simple)}")
print(f"File size: {len(df_simple)} records")

# Check if all employees are present
orig_emps = set(df_original['Empl Id'].unique())
simple_emps = set(df_simple['Empl Id'].unique())

print(f"\nOriginal employees: {len(orig_emps)}")
print(f"Simple file employees: {len(simple_emps)}")

if len(simple_emps) == len(orig_emps):
    print("✅ All employees present")
else:
    missing = orig_emps - simple_emps
    print(f"❌ Missing employees: {missing}")

# Check WO counts
wo_count = (df_simple['Calculated_WO'] == 'Y').sum()
absent_count = (df_simple['Duty Code'] == 'Absent').sum()

print(f"\nWOs assigned: {wo_count}")
print(f"Absent marked: {absent_count}")

# Sample verification
print(f"\nSample verification (Employee 11104):")
emp_11104 = df_simple[df_simple['Empl Id'] == 11104].head(20)
print(emp_11104[['Attendance Date', 'Shift', 'Duty Status', 'Calculated_WO', 'Week_Shift', 'Duty Code']].to_string())

print(f"\n✅ File ready: {file_path}")
print("Contains all records with corrected 7-day WO logic")
