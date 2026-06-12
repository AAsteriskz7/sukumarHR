#!/usr/bin/env python3
"""Explore Audit Work March 50.xlsx structure."""

import pandas as pd

file_path = "/Users/smartmysql/Documents/sukumarHR/Audit Work March 50.xlsx"

print(f"Reading file: {file_path}")

# Check sheet names
xl = pd.ExcelFile(file_path)
print(f"Sheet names: {xl.sheet_names}")

# Read each sheet
for sheet in xl.sheet_names:
    print(f"\n{'='*80}")
    print(f"SHEET: {sheet}")
    print('='*80)
    
    # Try different headers
    for header_row in range(3):
        try:
            df = pd.read_excel(file_path, sheet_name=sheet, header=header_row)
            if df.shape[0] > 0 and len(df.columns) > 3:
                print(f"\nHeader row {header_row}:")
                print(f"Columns: {df.columns.tolist()}")
                print(f"Shape: {df.shape}")
                print(f"\nFirst 10 rows:")
                print(df.head(10).to_string())
                
                if 'Empl Id' in df.columns or 'Employee' in str(df.columns):
                    print(f"\nUnique employees: {df['Empl Id'].nunique() if 'Empl Id' in df.columns else 'N/A'}")
                break
        except Exception as e:
            print(f"Error with header {header_row}: {e}")
            continue
