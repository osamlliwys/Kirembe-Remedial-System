#!/usr/bin/env python3
"""Analyze how to identify which stream (A or C) each learner belongs to."""

import openpyxl

wb = openpyxl.load_workbook('REMEDIAL_PAYMENT_BALANCES_2026.3_FIXED(1).xlsx')

print("=" * 80)
print("ANALYZING STREAM SECTIONS IN EACH SHEET")
print("=" * 80)

# Check GRADE 10 2026
sheet = wb['GRADE 10 2026']
print("\n\nGRADE 10 2026:")
print("-" * 80)
print("Looking for section dividers...")

for row in range(35, 110):
    a_val = sheet[f'A{row}'].value
    c_val = sheet[f'C{row}'].value
    
    if a_val and isinstance(a_val, str):
        if any(keyword in a_val.upper() for keyword in ['REMEDIAL', 'FEE REGISTER', 'GRADE 10']):
            print(f"  Row {row}: {a_val}")
    
    if c_val and isinstance(c_val, str) and 'TOTAL' in c_val.upper():
        print(f"  Row {row}: {c_val} (TOTAL marker)")

# Based on the output, determine the row ranges
print("\n\nDetermined Row Ranges:")
print("-" * 80)
print("GRADE 10 2026:")
print("  Grade 10A: Rows 40-58 (before row 59 TOTAL)")
print("  Grade 10C: Rows 72-100 (before row 101 TOTAL)")

# Check other sheets
for sheet_name in ['FORM 2 2025', 'FORM 3 2026', 'FORM 3 2025', 'FORM 4 2026']:
    sheet = wb[sheet_name]
    print(f"\n{sheet_name}:")
    
    for row in range(35, 110):
        a_val = sheet[f'A{row}'].value
        c_val = sheet[f'C{row}'].value
        
        if a_val and isinstance(a_val, str):
            if any(keyword in a_val.upper() for keyword in ['REMEDIAL', 'FEE REGISTER', 'FORM']):
                print(f"  Row {row}: {a_val}")
        
        if c_val and isinstance(c_val, str) and 'TOTAL' in c_val.upper():
            print(f"  Row {row}: {c_val}")

wb.close()
