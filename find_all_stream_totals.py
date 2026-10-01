#!/usr/bin/env python3
"""Find all stream total cells across sheets."""

import openpyxl

wb = openpyxl.load_workbook('REMEDIAL_PAYMENT_BALANCES_2026.3_FIXED(1).xlsx')

print("=" * 80)
print("STREAM TOTALS MAPPING")
print("=" * 80)

# Define the mapping based on user's description and our analysis
stream_mapping = {
    'GRADE 10 2026': {
        'Grade 10A 2026': 'H59',
        'Grade 10C 2026': 'E103',  # This is H101-H59
    },
    'FORM 3 2026': {
        'Form 3 2026': 'D79',  # User mentioned this
    },
    'FORM 3 2025': {
        'Form 3A 2025': 'F69',  # Based on "TOTAL COLLECTION 2025 (3A)"
    },
    'FORM 4 2026': {
        'Form 4 2026': 'F68',  # Similar structure to Form 3
    },
    'FORM 2 2025': {
        'Form 2 2025': 'D79',
    }
}

print("\nProposed Stream-to-Cell Mapping:")
print("-" * 80)

for sheet_name, streams in stream_mapping.items():
    print(f"\n{sheet_name}:")
    sheet = wb[sheet_name]
    for stream_name, cell_ref in streams.items():
        cell_value = sheet[cell_ref].value
        print(f"  {stream_name}: {cell_ref} = {cell_value}")

# Now let's look for more streams by checking for section headers
print("\n\n" + "=" * 80)
print("SEARCHING FOR ADDITIONAL STREAMS")
print("=" * 80)

for sheet_name in wb.sheetnames:
    sheet = wb[sheet_name]
    print(f"\n{sheet_name}:")
    
    # Look for rows with "TOTAL" or stream indicators
    for row in range(35, 110):
        c_val = sheet[f'C{row}'].value
        a_val = sheet[f'A{row}'].value
        e_val = sheet[f'E{row}'].value
        
        if c_val and isinstance(c_val, str) and 'TOTAL' in c_val.upper():
            h_val = sheet[f'H{row}'].value
            print(f"  Row {row}: C='{c_val}', H={h_val}")
        
        if a_val and isinstance(a_val, str):
            if any(keyword in a_val.upper() for keyword in ['FEE REGISTER', 'GRADE', 'FORM']):
                print(f"  Row {row}: A='{a_val}'")
        
        if e_val and isinstance(e_val, str):
            if 'TOTAL' in e_val.upper() or any(stream in e_val.upper() for stream in ['3A', '3C', '4A', '4C']):
                print(f"  Row {row}: E='{e_val}'")

wb.close()
