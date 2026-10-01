#!/usr/bin/env python3
"""
Calculate the formulas manually and update the Excel file.
This bypasses Excel's calculation engine and directly computes the values.
"""
import openpyxl
from decimal import Decimal
import shutil
from datetime import datetime

EXCEL_FILE = "REMEDIAL_PAYMENT_BALANCES_2026.3_FIXED(1).xlsx"

# Create backup first
backup_file = f"REMEDIAL_PAYMENT_BALANCES_2026.3_BACKUP_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
print(f"Creating backup: {backup_file}")
shutil.copy2(EXCEL_FILE, backup_file)
print("✓ Backup created")
print()

print("=" * 80)
print("CALCULATING FORMULAS IN FORM 3 2025")
print("=" * 80)
print()

# Load workbook
wb = openpyxl.load_workbook(EXCEL_FILE)
sheet_form3_2025 = wb['FORM 3 2025']

# Calculate column AD for each row
# Formula: =F{row}+SUM(G{row}:Q{row})-SUM(R{row}:AC{row})
# This means: ARREAS + (DEBT columns) - (REPAYMENT columns)

rows_updated = 0
for row in range(5, sheet_form3_2025.max_row + 1):
    # Check if this row has an admission number
    adm = sheet_form3_2025[f'D{row}'].value
    if not adm or adm == 'ADM':  # Skip header and empty rows
        continue
    
    # Get F value (ARREAS)
    f_val = sheet_form3_2025[f'F{row}'].value
    if f_val is None:
        f_val = 0
    
    # Sum G through Q (DEBT columns)
    debt_sum = 0
    for col_letter in ['G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q']:
        val = sheet_form3_2025[f'{col_letter}{row}'].value
        if val is not None and isinstance(val, (int, float)):
            debt_sum += val
    
    # Sum R through AC (REPAYMENT columns)
    repayment_sum = 0
    for col_idx in range(18, 30):  # R=18, AC=29
        col_letter = openpyxl.utils.get_column_letter(col_idx)
        val = sheet_form3_2025[f'{col_letter}{row}'].value
        if val is not None and isinstance(val, (int, float)):
            repayment_sum += val
    
    # Calculate: ARREAS + DEBT - REPAYMENTS
    calculated_value = float(f_val) + debt_sum - repayment_sum
    
    # Update cell AD
    sheet_form3_2025[f'AD{row}'].value = calculated_value
    rows_updated += 1
    
    name = sheet_form3_2025[f'E{row}'].value
    print(f"Row {row} - ADM {adm} ({name}): AD = {calculated_value:.2f}")

print()
print(f"✓ Updated {rows_updated} rows in FORM 3 2025")
print()

print("=" * 80)
print("CALCULATING FORMULAS IN FORM 4 2026")
print("=" * 80)
print()

sheet_form4_2026 = wb['FORM 4 2026']

# Now update FORM 4 2026 Column F to reference the calculated values
rows_updated = 0
for row in range(5, sheet_form4_2026.max_row + 1):
    # Check if this row has an admission number
    adm = sheet_form4_2026[f'D{row}'].value
    if not adm or adm == 'ADM':  # Skip header and empty rows
        continue
    
    # Check if column F has a formula referencing FORM 3 2025
    f_cell = sheet_form4_2026[f'F{row}']
    if isinstance(f_cell.value, str) and f_cell.value.startswith("='FORM 3 2025'!"):
        # Extract the cell reference (e.g., AD30 from ='FORM 3 2025'!AD30)
        ref = f_cell.value.split('!')[-1]
        
        # Get the value from FORM 3 2025
        referenced_value = sheet_form3_2025[ref].value
        
        # Update FORM 4 2026 with the calculated value
        sheet_form4_2026[f'F{row}'].value = referenced_value
        rows_updated += 1
        
        name = sheet_form4_2026[f'E{row}'].value
        g_val = sheet_form4_2026[f'G{row}'].value or 0
        ref_val = referenced_value if referenced_value is not None else 0
        total = float(ref_val) + float(g_val)
        
        print(f"Row {row} - ADM {adm} ({name}):")
        print(f"  F (2024 arrears): {ref_val:.2f}")
        print(f"  G (2025 debt): {g_val:.2f}")
        print(f"  Total owed: {total:.2f}")

print()
print(f"✓ Updated {rows_updated} rows in FORM 4 2026")
print()

# Save the workbook
print("Saving workbook...")
wb.save(EXCEL_FILE)
print("✓ Workbook saved")
print()

print("=" * 80)
print("SUCCESS!")
print("=" * 80)
print()
print("All formulas have been calculated and the Excel file has been updated.")
print()
print("Next step: Restart the web app to see the corrected balances.")
print()
print(f"If anything goes wrong, restore from backup: {backup_file}")

wb.close()
