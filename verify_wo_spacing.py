#!/usr/bin/env python3
"""Verify every WO has exactly 6 days before it."""

import pandas as pd

file_path = "/Users/smartmysql/Documents/sukumarHR/Audit Work March 50_7day.xlsx"
df = pd.read_excel(file_path, sheet_name='Processed_Data')

print("Verifying WO spacing for all employees...")
print("="*80)

errors_found = 0
total_wo_count = 0
employees_with_errors = []

for emp_id in df['Empl Id'].unique():
    emp_data = df[df['Empl Id'] == emp_id].copy()
    emp_data = emp_data.sort_values('Attendance Date').reset_index(drop=True)
    
    # Find all WOs (Calculated_WO = 'Y')
    wo_positions = []
    for idx, row in emp_data.iterrows():
        if row['Calculated_WO'] == 'Y':
            wo_positions.append(idx)
    
    if not wo_positions:
        continue
    
    total_wo_count += len(wo_positions)
    
    # Check spacing between WOs
    for i in range(1, len(wo_positions)):
        prev_wo = wo_positions[i-1]
        curr_wo = wo_positions[i]
        
        # Count days between WOs
        days_between = curr_wo - prev_wo
        
        # Check if exactly 7 days (6 working days + WO)
        if days_between != 7:
            errors_found += 1
            employees_with_errors.append(emp_id)
            
            print(f"\nERROR - Employee {emp_id}:")
            print(f"  WO at row {prev_wo+1}: {emp_data.loc[prev_wo, 'Attendance Date'].strftime('%Y-%m-%d')}")
            print(f"  WO at row {curr_wo+1}: {emp_data.loc[curr_wo, 'Attendance Date'].strftime('%Y-%m-%d')}")
            print(f"  Days between: {days_between} (should be 7)")
            
            # Show the 6 days between WOs
            print(f"  Days between WOs:")
            for j in range(prev_wo + 1, curr_wo):
                date = emp_data.loc[j, 'Attendance Date']
                shift = emp_data.loc[j, 'Week_Shift']
                duty_code = emp_data.loc[j, 'Duty Code']
                print(f"    Row {j+1}: {date.strftime('%Y-%m-%d')} - Shift: {shift}, Duty: {duty_code}")
            
            # Limit output to first 10 errors
            if errors_found >= 10:
                print("\n... (showing first 10 errors only)")
                break

if errors_found == 0:
    print(f"✅ SUCCESS: All {total_wo_count} WOs have exactly 6 days between them!")
else:
    print(f"\n❌ ERRORS FOUND: {errors_found} WOs with incorrect spacing")
    print(f"Total employees with errors: {len(set(employees_with_errors))}")
    print(f"Unique employees with errors: {set(employees_with_errors)}")

# Additional check: Verify each week has exactly 6 non-WO days
print("\n" + "="*80)
print("Additional check: Weekly structure verification")
print("="*80)

week_errors = 0
for emp_id in df['Empl Id'].unique()[:5]:  # Check first 5 employees as sample
    emp_data = df[df['Empl Id'] == emp_id].copy()
    emp_data = emp_data.sort_values('Attendance Date').reset_index(drop=True)
    
    print(f"\nEmployee {emp_id} weekly structure:")
    
    current_week = []
    for idx, row in emp_data.iterrows():
        if row['Calculated_WO'] == 'Y':
            if current_week:
                print(f"  Week before WO: {len(current_week)} days")
                if len(current_week) != 6:
                    week_errors += 1
                    print(f"    ERROR: Should be 6 days, found {len(current_week)}")
            current_week = []
        else:
            current_week.append(idx)
    
    # Check final week
    if current_week:
        print(f"  Final week: {len(current_week)} days")

if week_errors == 0:
    print("\n✅ Weekly structure check passed for sample employees!")
else:
    print(f"\n❌ Weekly structure errors: {week_errors}")
