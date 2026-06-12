#!/usr/bin/env python3
"""Quick verification of simple 7-day logic."""

import pandas as pd

file_path = "/Users/smartmysql/Documents/sukumarHR/Audit Work March 50_7day_simple.xlsx"
df = pd.read_excel(file_path, sheet_name='Processed_Data')

print("Verifying simple 7-day logic...")
print("="*60)

total_errors = 0
sample_size = 0

for emp_id in df['Empl Id'].unique()[:5]:  # Check first 5 employees
    emp_data = df[df['Empl Id'] == emp_id].copy()
    emp_data = emp_data.sort_values('Attendance Date').reset_index(drop=True)
    
    wo_positions = []
    for idx, row in emp_data.iterrows():
        if row['Calculated_WO'] == 'Y':
            wo_positions.append(idx)
    
    if not wo_positions:
        continue
    
    sample_size += 1
    print(f"\nEmployee {emp_id}: {len(wo_positions)} WOs")
    
    for i in range(1, len(wo_positions)):
        prev_wo = wo_positions[i-1]
        curr_wo = wo_positions[i]
        working_days = curr_wo - prev_wo - 1
        
        if working_days != 6:
            total_errors += 1
            print(f"  ERROR Week {i}: {working_days} working days (should be 6)")
            print(f"    From row {prev_wo+2} to {curr_wo+2}")

print(f"\n{'='*60}")
if total_errors == 0:
    print("✅ SUCCESS: All weeks have exactly 6 working days!")
else:
    print(f"❌ {total_errors} errors in {sample_size} sample employees")

# Overall stats
wo_count = (df['Calculated_WO'] == 'Y').sum()
print(f"\nTotal WOs assigned: {wo_count}")
print("File: Audit Work March 50_7day_simple.xlsx")
