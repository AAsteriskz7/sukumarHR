#!/usr/bin/env python3
"""Fixed test: Verify every week has exactly 6 working days between WOs."""

import pandas as pd

file_path = "/Users/smartmysql/Documents/sukumarHR/Audit Work March 50_7day.xlsx"
df = pd.read_excel(file_path, sheet_name='Processed_Data')

print("FIXED TEST: Weekly Structure Validation")
print("="*80)
print("Checking that every week has exactly 6 working days between WOs")
print("Pattern: WO -> 6 working days -> WO")
print("="*80)

total_errors = 0
error_details = []
employees_tested = 0

for emp_id in df['Empl Id'].unique():
    emp_data = df[df['Empl Id'] == emp_id].copy()
    emp_data = emp_data.sort_values('Attendance Date').reset_index(drop=True)
    
    # Find all WOs for this employee
    wo_positions = []
    for idx, row in emp_data.iterrows():
        if row['Calculated_WO'] == 'Y':
            wo_positions.append(idx)
    
    if not wo_positions:
        continue
    
    employees_tested += 1
    emp_errors = 0
    
    # Check each week between WOs
    for i in range(1, len(wo_positions)):
        prev_wo = wo_positions[i-1]
        curr_wo = wo_positions[i]
        
        # Count working days between WOs (exclusive of both WOs)
        working_days = curr_wo - prev_wo - 1
        
        if working_days != 6:
            total_errors += 1
            emp_errors += 1
            
            error_details.append({
                'emp_id': emp_id,
                'week_num': i,
                'prev_wo_row': prev_wo + 2,  # Excel row number
                'curr_wo_row': curr_wo + 2,
                'prev_wo_date': emp_data.loc[prev_wo, 'Attendance Date'],
                'curr_wo_date': emp_data.loc[curr_wo, 'Attendance Date'],
                'working_days': working_days
            })
            
            # Show details for first few errors
            if total_errors <= 10:
                print(f"\nERROR #{total_errors}: Employee {emp_id}, Week {i}")
                print(f"  Previous WO: Row {prev_wo+2}, {emp_data.loc[prev_wo, 'Attendance Date'].strftime('%Y-%m-%d')}")
                print(f"  Current WO:  Row {curr_wo+2}, {emp_data.loc[curr_wo, 'Attendance Date'].strftime('%Y-%m-%d')}")
                print(f"  Working days between: {working_days} (should be 6)")
                
                # Show the working days in this week
                print(f"  Working days between WOs:")
                for j in range(prev_wo + 1, curr_wo):
                    date = emp_data.loc[j, 'Attendance Date']
                    duty_code = emp_data.loc[j, 'Duty Code']
                    week_shift = emp_data.loc[j, 'Week_Shift']
                    
                    if duty_code == 'Absent':
                        day_type = "Absent"
                    elif week_shift == '0':
                        day_type = "Shift_0"
                    elif week_shift == 'WO':
                        day_type = "WO_Error"
                    else:
                        day_type = f"Shift_{week_shift}"
                    
                    print(f"    Row {j+2:3d}: {date.strftime('%Y-%m-%d')} - {day_type}")

# Summary
print("\n" + "="*80)
print("TEST SUMMARY")
print("="*80)
print(f"Total employees tested: {employees_tested}")
print(f"Total errors found: {total_errors}")

if total_errors == 0:
    print("✅ SUCCESS: All weeks have exactly 6 working days between WOs!")
else:
    print(f"❌ ERRORS: {total_errors} weeks don't have exactly 6 working days")
    
    # Analyze error patterns
    less_than_6 = sum(1 for e in error_details if e['working_days'] < 6)
    more_than_6 = sum(1 for e in error_details if e['working_days'] > 6)
    
    print(f"\nError breakdown:")
    print(f"  Weeks with < 6 working days: {less_than_6}")
    print(f"  Weeks with > 6 working days: {more_than_6}")
    
    # Show unique day counts
    day_counts = {}
    for e in error_details:
        days = e['working_days']
        day_counts[days] = day_counts.get(days, 0) + 1
    
    print(f"\nWorking day count distribution:")
    for days in sorted(day_counts.keys()):
        print(f"  {days} working days: {day_counts[days]} occurrences")
    
    # Show affected employees
    affected_employees = set(e['emp_id'] for e in error_details)
    print(f"\nAffected employees: {len(affected_employees)} out of {employees_tested}")

print("\n" + "="*80)
print("EXPECTED RESULT: WO -> 6 working days -> WO")
print("- Total cycle: 7 days (1 WO + 6 working days)")
print("- Working days can be shifts, 0, SL, HO, Absent, etc.")
print("="*80)
