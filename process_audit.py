#!/usr/bin/env python3
"""
Script to process Audit Work March 46.xlsx and add new columns for WO allocation.
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import os

# Read the Excel file - check all sheets
file_path = "/Users/smartmysql/Documents/sukumarHR/Audit Work March 46.xlsx"
output_path = "/Users/smartmysql/Documents/sukumarHR/Audit Work March 46_processed.xlsx"

print(f"Reading file: {file_path}")

# Check sheet names
xl = pd.ExcelFile(file_path)
print(f"Sheet names: {xl.sheet_names}")

# Read Sheet2 - likely contains the actual employee data
for sheet in xl.sheet_names:
    print(f"\n\n========== SHEET: {sheet} ==========")
    for header_row in range(5):
        try:
            df = pd.read_excel(file_path, sheet_name=sheet, header=header_row)
            print(f"\n--- Header row {header_row} ---")
            print(f"Columns: {df.columns.tolist()}")
            print(f"Shape: {df.shape}")
            if df.shape[0] > 0:
                print(df.head(3).to_string())
                break
        except Exception as e:
            print(f"Error with header {header_row}: {e}")
            continue
