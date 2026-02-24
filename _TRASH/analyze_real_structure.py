#!/usr/bin/env python3
"""
REAL ANALYSIS: What are the ACTUAL column labels and formulas in DSF?
Let's verify what the columns REALLY contain before we fill anything.
"""

from openpyxl import load_workbook
from pathlib import Path

print("=" * 80)
print("EXAMINING REAL DSF STRUCTURE")
print("=" * 80)

template_file = Path('templates/DSF Normal standard.xlsx')
balance_file = Path('input/BALANCE AU 31 12 2024 VERSION DU 26 05 25.xlsx')

# Check DSF template column headers
print("\n\n1. DSF TEMPLATE - Column headers (rows 8-10, first 15 columns)")
print("-" * 80)

wb_dsf = load_workbook(template_file)
sheets_to_check = ['NOTE 3A', 'NOTE 4', 'NOTE 16A']  # Check high-volume sheets

for sheet_name in sheets_to_check:
    if sheet_name not in wb_dsf.sheetnames:
        continue
    
    ws = wb_dsf[sheet_name]
    print(f"\n{sheet_name}:")
    print("  Row 8 (section headers):")
    for col in range(1, 16):
        cell = ws.cell(row=8, column=col)
        if cell.value:
            print(f"    Col {col}: {cell.value}")
    
    print("  Row 9 (sub-headers):")
    for col in range(1, 16):
        cell = ws.cell(row=9, column=col)
        if cell.value:
            print(f"    Col {col}: {cell.value}")
    
    print("  Row 10 (detail headers):")
    for col in range(1, 16):
        cell = ws.cell(row=10, column=col)
        if cell.value:
            print(f"    Col {col}: {cell.value}")
    
    # Check row 11+ for formulas
    print("  Row 11+ - Any formulas?")
    found_formula = False
    for row in range(11, 20):
        for col in range(1, 10):
            cell = ws.cell(row=row, column=col)
            if cell.data_type == 'f':  # Formula
                print(f"    {cell.coordinate}: {cell.value}")
                found_formula = True
    if not found_formula:
        print("    (No formulas found in rows 11-19)")

# Check balance structure
print("\n\n2. BALANCE FILE - What columns actually exist?")
print("-" * 80)

wb_balance = load_workbook(balance_file)
ws_balance = wb_balance.active

print(f"Sheet: {ws_balance.title}")
print(f"Rows: {ws_balance.max_row}, Columns: {ws_balance.max_column}")

print("\nHeaders (rows 1-10):")
for row in range(1, 11):
    values = []
    for col in range(1, 30):
        cell = ws_balance.cell(row=row, column=col)
        if cell.value:
            values.append(f"C{col}:{cell.value}")
    if values:
        print(f"  Row {row}: {', '.join(values[:5])}")

print("\nData sample (row 25):")
for col in range(1, 30):
    cell = ws_balance.cell(row=25, column=col)
    if cell.value and col <= 15:
        print(f"  Col {col}: {cell.value}")

# Check if DSF columns reference balance columns
print("\n\n3. VERIFICATION: Do DSF cells contain formulas that reference balance?")
print("-" * 80)

for sheet_name in sheets_to_check:
    if sheet_name not in wb_dsf.sheetnames:
        continue
    
    ws = wb_dsf[sheet_name]
    print(f"\n{sheet_name} - Sampling formulas:")
    formula_count = 0
    
    for row in range(11, 50):
        for col in range(5, 15):
            cell = ws.cell(row=row, column=col)
            if cell.data_type == 'f':  # Formula
                formula_count += 1
                if formula_count <= 5:  # Print first 5
                    print(f"  {cell.coordinate}: {cell.value}")
    
    if formula_count == 0:
        print("  NO FORMULAS FOUND - cells are all EMPTY")
    else:
        print(f"  Total formulas in sample: {formula_count}")

print("\n" + "=" * 80)
print("CONCLUSION")
print("=" * 80)
print("""
If DSF cells are EMPTY with NO formulas:
  → We should fill them with values from balance

If DSF cells have FORMULAS:
  → We should NOT fill them (formulas will calculate)
  → We should fill only the INPUTS they reference

We need to understand:
  - Which columns are DEBIT vs CREDIT in BOTH files
  - Whether column labels give us MAPPING HINTS
  - Whether we're filling OUTPUT cells or INPUT cells
""")
