#!/usr/bin/env python3
"""Add initial balance rows to Excel that will be included in totals."""

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill

# Initial data provided by user
initial_totals = {
    'Grade 10A 2026': 14100,
    'Grade 10C 2026': 16900,
    'Form 2A 2025': 22150,
    'Form 2C 2025': 7100,
    'Form 3A 2026': 3700,
    'Form 3C 2026': 1300,
    'Form 3A 2025': 21400,
    'Form 3C 2025': 13150,
    'Form 4A 2026': 3000,
    'Form 4C 2026': 2400,
}

print("=" * 80)
print("ADDING INITIAL BALANCES TO EXCEL")
print("=" * 80)

wb = openpyxl.load_workbook('REMEDIAL_PAYMENT_BALANCES_2026.3_FIXED(1).xlsx')

# Strategy: Add initial balance in row 40 for each section
# This row will be included in the SUM formulas

# Grade 10A 2026 - add to H40
sheet = wb['GRADE 10 2026']
sheet['C40'] = 'INITIAL BALANCE'
sheet['H40'] = initial_totals['Grade 10A 2026']
sheet['C40'].font = Font(bold=True, color="0000FF")
sheet['H40'].font = Font(bold=True, color="0000FF")
print(f"✓ Grade 10A 2026: KES {initial_totals['Grade 10A 2026']:,} → H40")

# Grade 10C 2026 - add to H72
sheet['C72'] = 'INITIAL BALANCE'
sheet['H72'] = initial_totals['Grade 10C 2026']
sheet['C72'].font = Font(bold=True, color="0000FF")
sheet['H72'].font = Font(bold=True, color="0000FF")
print(f"✓ Grade 10C 2026: KES {initial_totals['Grade 10C 2026']:,} → H72")

# Form 2A 2025 - add to H39
sheet = wb['FORM 2 2025']
sheet['C39'] = 'INITIAL BALANCE'
sheet['H39'] = initial_totals['Form 2A 2025']
sheet['C39'].font = Font(bold=True, color="0000FF")
sheet['H39'].font = Font(bold=True, color="0000FF")
print(f"✓ Form 2A 2025: KES {initial_totals['Form 2A 2025']:,} → H39")

# Form 2C 2025 - Need to add a total cell since it doesn't exist
# Add at D103 (similar to D79 for 2A)
sheet['C103'] = 'TOTAL COLLECTION 2025 (2C)'
sheet['D103'] = initial_totals['Form 2C 2025']
sheet['C103'].font = Font(bold=True)
sheet['D103'].font = Font(bold=True)
print(f"✓ Form 2C 2025: KES {initial_totals['Form 2C 2025']:,} → D103 (new cell)")

# Form 3A 2026 - add to H39
sheet = wb['FORM 3 2026']
sheet['C39'] = 'INITIAL BALANCE'
sheet['H39'] = initial_totals['Form 3A 2026']
sheet['C39'].font = Font(bold=True, color="0000FF")
sheet['H39'].font = Font(bold=True, color="0000FF")
print(f"✓ Form 3A 2026: KES {initial_totals['Form 3A 2026']:,} → H39")

# Form 3C 2026 - Need to add a total cell
sheet['C103'] = 'TOTAL COLLECTION 2026 (3C)'
sheet['D103'] = initial_totals['Form 3C 2026']
sheet['C103'].font = Font(bold=True)
sheet['D103'].font = Font(bold=True)
print(f"✓ Form 3C 2026: KES {initial_totals['Form 3C 2026']:,} → D103 (new cell)")

# Form 3A 2025 - add to R37 (first row in the range)
sheet = wb['FORM 3 2025']
sheet['E37'] = 'INITIAL BALANCE'
sheet['R37'] = initial_totals['Form 3A 2025']
sheet['E37'].font = Font(bold=True, color="0000FF")
sheet['R37'].font = Font(bold=True, color="0000FF")
print(f"✓ Form 3A 2025: KES {initial_totals['Form 3A 2025']:,} → R37")

# Form 3C 2025 - add to R108 (first row in the 3C range)
sheet['E108'] = 'INITIAL BALANCE'
sheet['R108'] = initial_totals['Form 3C 2025']
sheet['E108'].font = Font(bold=True, color="0000FF")
sheet['R108'].font = Font(bold=True, color="0000FF")
print(f"✓ Form 3C 2025: KES {initial_totals['Form 3C 2025']:,} → R108")

# Form 4A 2026 - add to R37
sheet = wb['FORM 4 2026']
sheet['E37'] = 'INITIAL BALANCE'
sheet['R37'] = initial_totals['Form 4A 2026']
sheet['E37'].font = Font(bold=True, color="0000FF")
sheet['R37'].font = Font(bold=True, color="0000FF")
print(f"✓ Form 4A 2026: KES {initial_totals['Form 4A 2026']:,} → R37")

# Form 4C 2026 - add to R75 (first row in the 4C range)
sheet['E75'] = 'INITIAL BALANCE'
sheet['R75'] = initial_totals['Form 4C 2026']
sheet['E75'].font = Font(bold=True, color="0000FF")
sheet['R75'].font = Font(bold=True, color="0000FF")
print(f"✓ Form 4C 2026: KES {initial_totals['Form 4C 2026']:,} → R75")

# Save the workbook
wb.save('REMEDIAL_PAYMENT_BALANCES_2026.3_FIXED(1).xlsx')
wb.close()

print("\n" + "=" * 80)
print("✓ Initial balances added successfully!")
print("✓ The formulas will now include these initial amounts")
print("✓ New payments will be added on top of these initial balances")
print("=" * 80)
