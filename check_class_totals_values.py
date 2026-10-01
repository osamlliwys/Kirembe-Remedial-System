#!/usr/bin/env python3
"""Script to check actual values in class total cells."""

import openpyxl

wb = openpyxl.load_workbook('REMEDIAL_PAYMENT_BALANCES_2026.3_FIXED(1).xlsx', data_only=True)

print("=" * 80)
print("CLASS TOTALS - ACTUAL VALUES")
print("=" * 80)

# Grade 10 2026
sheet = wb['GRADE 10 2026']
h59_val = sheet['H59'].value or 0
h101_val = sheet['H101'].value or 0
e103_val = sheet['E103'].value or 0
print(f"\nGRADE 10 2026:")
print(f"  H59 (Grade 10A total): {h59_val}")
print(f"  H101 (total): {h101_val}")
print(f"  E103 (H101-H59, Grade 10C): {e103_val}")
print(f"  Calculated H101-H59: {h101_val - h59_val}")

# Form 3 2026
sheet = wb['FORM 3 2026']
d79_val = sheet['D79'].value or 0
h76_val = sheet['H76'].value or 0
print(f"\nFORM 3 2026:")
print(f"  D79 (total): {d79_val}")
print(f"  H76 (H column sum): {h76_val}")

# Form 3 2025
sheet = wb['FORM 3 2025']
f69_val = sheet['F69'].value or 0
f67_val = sheet['F67'].value or 0
h67_val = sheet['H67'].value or 0
print(f"\nFORM 3 2025:")
print(f"  F69 (total): {f69_val}")
print(f"  F67 (F column sum): {f67_val}")
print(f"  H67 (H column sum): {h67_val}")

# Form 4 2026
sheet = wb['FORM 4 2026']
f68_val = sheet['F68'].value or 0
f67_val = sheet['F67'].value or 0
h67_val = sheet['H67'].value or 0
print(f"\nFORM 4 2026:")
print(f"  F68 (total): {f68_val}")
print(f"  F67 (F column sum): {f67_val}")
print(f"  H67 (H column sum): {h67_val}")

# Form 2 2025
sheet = wb['FORM 2 2025']
d79_val = sheet['D79'].value or 0
h76_val = sheet['H76'].value or 0
print(f"\nFORM 2 2025:")
print(f"  D79 (total): {d79_val}")
print(f"  H76 (H column sum): {h76_val}")

wb.close()
