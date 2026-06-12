#!/usr/bin/env python3
"""Import Excel data to SQLite for analysis."""

import pandas as pd
import sqlite3
import os

# File paths
excel_file = "/Users/smartmysql/Documents/sukumarHR/Audit Work March 50_7day_simple.xlsx"
sqlite_file = "/Users/smartmysql/Documents/sukumarHR/audit_data.sqlite"

print("Importing data to SQLite...")
print(f"Excel file: {excel_file}")
print(f"SQLite file: {sqlite_file}")

# Remove existing SQLite file if it exists
if os.path.exists(sqlite_file):
    os.remove(sqlite_file)

# Connect to SQLite
conn = sqlite3.connect(sqlite_file)

# Read Excel data
df = pd.read_excel(excel_file, sheet_name='Processed_Data')

# Clean column names (remove spaces and special characters)
df.columns = [col.replace(' ', '_').replace('.', '_') for col in df.columns]

# Convert dates
df['Attendance_Date'] = pd.to_datetime(df['Attendance_Date'])

# Create table
df.to_sql('attendance', conn, if_exists='replace', index=False)

print(f"✅ Imported {len(df)} records to SQLite")

# Create some useful indexes
cursor = conn.cursor()
cursor.execute("CREATE INDEX idx_emp_id ON attendance(Empl_Id)")
cursor.execute("CREATE INDEX idx_date ON attendance(Attendance_Date)")
cursor.execute("CREATE INDEX idx_wo ON attendance(Calculated_WO)")
cursor.execute("CREATE INDEX idx_emp_date ON attendance(Empl_Id, Attendance_Date)")

conn.commit()
conn.close()

print("✅ Created indexes for faster queries")
print(f"\nDatabase ready: {sqlite_file}")
print("\nSample queries:")
print("1. Check WO spacing:")
print("   SELECT Empl_Id, COUNT(*) as wo_count FROM attendance WHERE Calculated_WO = 'Y' GROUP BY Empl_Id")
print("\n2. Check weekly structure:")
print("   SELECT Empl_Id, Attendance_Date, Calculated_WO FROM attendance WHERE Empl_Id = 11104 ORDER BY Attendance_Date")
