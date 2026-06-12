#!/usr/bin/env python3
"""Verify the processed file output."""

import pandas as pd

output_path = "/Users/smartmysql/Documents/sukumarHR/Audit Work March 46_processed.xlsx"

print(f"Reading processed file: {output_path}")
df = pd.read_excel(output_path, sheet_name='Processed_Data')

print(f"\nShape: {df.shape}")
print(f"Columns: {df.columns.tolist()}")

# Show summary of new columns
print("\n=== NEW COLUMNS SUMMARY ===")
print(f"Calculated_WO value counts:\n{df['Calculated_WO'].value_counts()}")
print(f"\nWO_Sequence range: {df['WO_Sequence'].min()} to {df['WO_Sequence'].max()}")
print(f"Days_Since_Last_WO range: {df['Days_Since_Last_WO'].min()} to {df['Days_Since_Last_WO'].max()}")

# Show unique calculated shifts
print(f"\nCalculated_Shift unique values: {sorted(df['Calculated_Shift'].dropna().unique())}")

# Show sample of first employee
print("\n=== FIRST EMPLOYEE DATA SAMPLE ===")
emp_id = df['Empl Id'].iloc[0]
sample = df[df['Empl Id'] == emp_id][['Empl Id', 'Attendance Date', 'Shift', 'Remarks', 'Duty Status', 
                                        'Calculated_WO', 'Calculated_Shift', 'WO_Sequence', 'Days_Since_Last_WO']].head(30)
print(sample.to_string())
