#!/usr/bin/env python3
"""Comprehensive test: Verify every week has exactly 7 days (6 working + 1 WO)."""

import pandas as pd

file_path = "/Users/smartmysql/Documents/sukumarHR/Audit Work March 50_7day.xlsx"
df = pd.read_excel(file_path, sheet_name='Processed_Data')

print("COMPREHENSIVE TEST: Weekly Structure Validation")
print("="*80)
print("Checking that every week has exactly 7 days (6 working + 1 WO)")
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
        
        # Calculate total days in week (inclusive of both WOs)
        total_days = curr_wo - prev_wo + 1
        
        if total_days != 7:
            total_errors += 1
            emp_errors += 1
            
            error_details.append({
                'emp_id': emp_id,
                'week_num': i,
                'prev_wo_row': prev_wo + 2,  # Excel row number
                'curr_wo_row': curr_wo + 2,
                'prev_wo_date': emp_data.loc[prev_wo, 'Attendance Date'],
                'curr_wo_date': emp_data.loc[curr_wo, 'Attendance Date'],
                'total_days': total_days,
                'working_days': total_days - 1
            })
            
            # Show details for first few errors
            if total_errors <= 10:
                print(f"\nERROR #{total_errors}: Employee {emp_id}, Week {i}")
                print(f"  Previous WO: Row {prev_wo+2}, {emp_data.loc[prev_wo, 'Attendance Date'].strftime('%Y-%m-%d')}")
                print(f"  Current WO:  Row {curr_wo+2}, {emp_data.loc[curr_wo, 'Attendance Date'].strftime('%Y-%m-%d')}")
                print(f"  Total days: {total_days} (should be 7)")
                print(f"  Working days: {total_days-1} (should be 6)")
                
                # Show the days in this week
                print(f"  Days in this week:")
                for j in range(prev_wo, curr_wo + 1):
                    date = emp_data.loc[j, 'Attendance Date']
                    calc_wo = emp_data.loc[j, 'Calculated_WO']
                    duty_code = emp_data.loc[j, 'Duty Code']
                    week_shift = emp_data.loc[j, 'Week_Shift']
                    
                    if calc_wo == 'Y':
                        day_type = "WO"
                    elif duty_code == 'Absent':
                        day_type = "Absent"
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
    print("✅ SUCCESS: All weeks have exactly 7 days!")
else:
    print(f"❌ ERRORS: {total_errors} weeks don't have exactly 7 days")
    
    # Analyze error patterns
    less_than_7 = sum(1 for e in error_details if e['total_days'] < 7)
    more_than_7 = sum(1 for e in error_details if e['total_days'] > 7)
    
    print(f"\nError breakdown:")
    print(f"  Weeks with < 7 days: {less_than_7}")
    print(f"  Weeks with > 7 days: {more_than_7}")
    
    # Show unique day counts
    day_counts = {}
    for e in error_details:
        days = e['total_days']
        day_counts[days] = day_counts.get(days, 0) + 1
    
    print(f"\nDay count distribution:")
    for days in sorted(day_counts.keys()):
        print(f"  {days} days: {day_counts[days]} occurrences")
    
    # Show affected employees
    affected_employees = set(e['emp_id'] for e in error_details)
    print(f"\nAffected employees: {len(affected_employees)} out of {employees_tested}")
    if len(affected_employees) <= 20:
        print(f"Employee IDs: {sorted(affected_employees)}")

print("\n" + "="*80)
print("EXPECTED RESULT: Every week should be exactly 7 days")
print("- 1 WO day")
print("- 6 working days (can be shifts or Absent)")
print("="*80)
