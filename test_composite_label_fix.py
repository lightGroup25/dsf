#!/usr/bin/env python3
"""
Test VERIFICATION: Check that we're NOT filling same values in each column
And that we're properly using BOTH row_label + col_label composite
"""

import sys
from pathlib import Path
sys.path.insert(0, str(Path('src')))

from openpyxl import load_workbook
from semantic_balance_filler import SemanticBalanceFiller
from dsf_inventory import DSFInventory

print("=" * 80)
print("TEST - COLUMN DIFFERENTIATION (CRITICAL BUG FIX)")
print("=" * 80)

template_path = Path('templates/DSF Normal standard.xlsx')
balance_path = Path('input/BALANCE AU 31 12 2024 VERSION DU 26 05 25.xlsx')
inventory_path = Path('data/dsf_inventory.json')

# Load
print("\nLoading inventory...")
inventory = DSFInventory.from_json(inventory_path)

# Create filler
print(f"🔧 Creating filler with COMPOSITE label fix...")
filler = SemanticBalanceFiller(
    str(template_path),
    str(balance_path),
    inventory,
    fuzzy_threshold=0.50,
    enable_hybrid=False,  # Test base functionality
)

# Load data
print(f"📥 Loading...")
filler.load()

# Fill
print(f"\n🔄 Filling with composite label matching...")
try:
    filled_count = filler.fill()
    print(f"✓ Filled: {filled_count} cells")
except Exception as e:
    print(f"❌ Error: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

# Save
output_path = Path('output/test_composite_label_fix.xlsx')
filler.save(output_path)
print(f"✓ Output: {output_path}")

# VERIFY: Check that same rows DON'T have same values in different columns
print(f"\n{'=' * 80}")
print("🔍 VERIFICATION: Are different columns getting different values?")
print(f"{'=' * 80}")

wb = load_workbook(output_path, data_only=True)

verification_found = False
for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]
    
    # Check each row (skip headers)
    for row_num in range(11, min(ws.max_row + 1, 200)):  # Check first 200 rows
        row_values = []
        non_empty_col_nums = []
        
        # Collect all values in this row
        for col_num in range(1, ws.max_column + 1):
            cell = ws.cell(row=row_num, column=col_num)
            if cell.value and isinstance(cell.value, (int, float)) and cell.value != 0:
                row_values.append(cell.value)
                non_empty_col_nums.append(col_num)
        
        # If multiple non-zero values in same row, check if they're different
        if len(row_values) > 1:
            if len(set(row_values)) == 1:  # All same
                # Get row and col labels
                row_label = ws.cell(row=row_num, column=1).value or ws.cell(row=row_num, column=2).value
                col_labels = [ws.cell(row=10, column=c).value for c in non_empty_col_nums]
                
                print(f"\n⚠️ POTENTIAL BUG in {sheet_name} row {row_num}:")
                print(f"   Row label: {row_label}")
                print(f"   All columns have SAME value: {row_values[0]}")
                print(f"   Columns: {col_labels}")
                verification_found = True
            else:
                # Different values - GOOD
                row_label = ws.cell(row=row_num, column=1).value or ws.cell(row=row_num, column=2).value
                if verification_found == False:  # Only print first good example
                    print(f"\n✅ GOOD: Different columns have different values")
                    print(f"   Row label: {row_label}")
                    print(f"   Column values: {row_values}")
                    verification_found = "good"

if verification_found == "good":
    print(f"\n✅ VERIFICATION PASSED: Column differentiation is working!")
elif verification_found:
    print(f"\n❌ VERIFICATION FAILED: Found rows with identical values in multiple columns")
else:
    print(f"\n⚠️ VERIFICATION INCOMPLETE: No multi-column rows found (might be sparse data)")

# Statistics
print(f"\n{'=' * 80}")
print("📊 STATISTICS")
print(f"{'=' * 80}")
print(f"Total filled: {filled_count} cells")
stats = filler.get_stats()
print(f"Assignment confidence distribution:")
print(f"  High (>80%):    {stats['high_confidence']}")
print(f"  Medium (50-80%): {stats['assignments'] - stats['high_confidence']}")

print(f"\n💡 Expected improvement:")
print(f"   V5 (row-only):      572 cells (11.2%)")
print(f"   V7 (row+col fix):   ? cells (should be 15-25%)")
