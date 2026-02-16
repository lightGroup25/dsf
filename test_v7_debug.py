#!/usr/bin/env python3
"""Test V7: Simple version (debug where it fails)"""

import sys
from pathlib import Path
sys.path.insert(0, str(Path('src')))

from semantic_balance_filler import SemanticBalanceFiller
from dsf_inventory import DSFInventory

print("TEST V7 - DEBUG")

try:
    print("\n1. Loading inventory...")
    inventory = DSFInventory.from_json(Path('data/dsf_inventory.json'))
    print("   OK")
    
    print("2. Creating filler...")
    filler = SemanticBalanceFiller(
        'templates/DSF Normal standard.xlsx',
        'input/BALANCE AU 31 12 2024 VERSION DU 26 05 25.xlsx',
        inventory,
        fuzzy_threshold=0.50,
    )
    print("   OK")
    
    print("3. Loading template and balance...")
    filler.load()
    print(f"   OK - {len(filler.balance_accounts)} accounts loaded")
    
    print("4. Filling cells...")
    filled_count = filler.fill()
    print(f"   OK - {filled_count} cells filled")
    
    print("5. Saving output...")
    output_path = Path('output/test_v7_composite_fix.xlsx')
    filler.save(output_path)
    print(f"   OK - {output_path}")
    
    print(f"\nRESULTS:")
    print(f"  V5: 572 cells (11.21%)")
    print(f"  V7: {filled_count} cells ({filled_count/5102*100:.2f}%)")
    
    if filled_count > 572:
        delta = filled_count - 572
        delta_pct = delta / 572 * 100
        print(f"  Delta: +{delta} cells (+{(filled_count/5102*100 - 11.21):.2f}%)")
        
except Exception as e:
    print(f"ERROR at step: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)
