#!/usr/bin/env python3
"""Test the composite label fix"""

import sys
from pathlib import Path
sys.path.insert(0, str(Path('src')))

from semantic_balance_filler import SemanticBalanceFiller
from dsf_inventory import DSFInventory

print("\n" + "=" * 80)
print("TEST V7: COMPOSITE LABEL FIX (row + col)")
print("=" * 80)

template_path = Path('templates/DSF Normal standard.xlsx')
balance_path = Path('input/BALANCE AU 31 12 2024 VERSION DU 26 05 25.xlsx')
inventory_path = Path('data/dsf_inventory.json')
output_path = Path('output/test_v7_composite_fix.xlsx')

try:
    print("\nLoading...")
    inventory = DSFInventory.from_json(inventory_path)
    
    print("Creating filler (composite label enabled)...")
    filler = SemanticBalanceFiller(
        str(template_path),
        str(balance_path),
        inventory,
        fuzzy_threshold=0.50,
        enable_hybrid=False,
    )
    
    print("Loading and filling...")
    filler.load()
    filled_count = filler.fill()
    
    print(f"RESULTS:")
    print(f"  Filled: {filled_count} cells")
    
    # Save
    filler.save(output_path)
    print(f"  Output: {output_path}")
    
    # Get stats
    stats = filler.get_stats()
    print(f"\nStatistics:")
    print(f"  Assignments: {stats['assignments']}")
    print(f"  Total amount: {stats['total_amount']/1e9:.1f}B XAF")
    print(f"  High confidence: {stats['high_confidence']}")
    
    print(f"\nCOMPARISON:")
    print(f"  V5 (row-only):       572 cells (11.21%)")
    print(f"  V7 (row+col fixed):  {filled_count} cells ({filled_count/5102*100:.2f}%)")
    
    if filled_count > 572:
        improvement = filled_count - 572
        improvement_pct = (improvement / 572 * 100)
        print(f"  IMPROVEMENT: +{improvement} cells (+{improvement_pct:.1f}%)")
    
except Exception as e:
    print(f"ERROR: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)
