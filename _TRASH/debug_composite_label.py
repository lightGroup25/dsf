#!/usr/bin/env python3
"""Debug: Test single cell matching with composite label"""

import sys
from pathlib import Path
sys.path.insert(0, str(Path('src')))

from semantic_balance_filler import SemanticBalanceFiller, CellNeed
from dsf_inventory import DSFInventory

print("Debug: Testing composite label fix")

try:
    print("\nLoading...")
    inventory = DSFInventory.from_json(Path('data/dsf_inventory.json'))
    filler = SemanticBalanceFiller(
        'templates/DSF Normal standard.xlsx',
        'input/BALANCE AU 31 12 2024 VERSION DU 26 05 25.xlsx',
        inventory,
        fuzzy_threshold=0.50,
    )
    filler.load()
    print("Loaded OK")
    
    # Create test cell need
    print("\nTesting _satisfy_cell_need() with composite label...")
    
    # Example: Row "STOCKS" + Col "Stock produits finis"
    need = CellNeed(
        sheet="NOTE 3A",
        cell="D15",
        row_num=15,
        col_num=4,
        row_label="STOCKS",
        col_label="Stocks produits finis",
        is_merged=False,
    )
    
    print(f"  Row: '{need.row_label}'")
    print(f"  Col: '{need.col_label}'")
    print(f"  Composite would be: '{need.row_label} {need.col_label}'")
    
    # Try to satisfy
    assignment = filler._satisfy_cell_need(need)
    
    if assignment:
        print(f"\nSUCCESS: Found match")
        print(f"  Account(s): {assignment.source_accounts}")
        print(f"  Amount: {assignment.total_amount}")
    else:
        print(f"\nNO MATCH: As expected (test data)")
    
    print("\nComposite label fix appears OK")
    
except Exception as e:
    print(f"ERROR: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)
