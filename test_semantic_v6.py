#!/usr/bin/env python3
"""
Test Semantic Balance Filler V6 - HYBRID OPTIMIZATION
Adds 3 complementary strategies to fuzzy matching:
1. Hierarchical (prefix matching)
2. Pattern (word presence)
3. Category (account grouping)
Target: 20-25% success rate
"""

import sys
from pathlib import Path
from decimal import Decimal

# Add src to path
sys.path.insert(0, str(Path('src')))

from semantic_balance_filler import SemanticBalanceFiller
from semantic_validators import SemanticFillerValidator
from dsf_inventory import DSFInventory


def main():
    print("=" * 80)
    print("🧪 TEST V6 - SEMANTIC BALANCE FILLER (HYBRID OPTIMIZATION)")
    print("=" * 80)
    
    template_path = Path('templates/DSF Normal standard.xlsx')
    balance_path = Path('input/BALANCE AU 31 12 2024 VERSION DU 26 05 25.xlsx')
    inventory_path = Path('data/dsf_inventory.json')
    
    print(f"\n✓ Template: {template_path}")
    print(f"✓ Balance: {balance_path}")
    print(f"✓ Inventory: {inventory_path}")
    
    # Load inventory
    print(f"\n📖 Loading inventory...")
    inventory = DSFInventory.from_json(inventory_path)
    print(f"✓ Inventory loaded: {len(inventory._fields_by_sheet)} sheets, {sum(len(fields) for fields in inventory._fields_by_sheet.values())} fields")
    
    # Initialize filler with hybrid optimization (to be added)
    print(f"\n🔧 Initializing Semantic Balance Filler V6 (Hybrid)...")
    print(f"   Strategy 1: Fuzzy matching (threshold: 0.50)")
    print(f"   Strategy 2: Hierarchical matching (prefix-based)")
    print(f"   Strategy 3: Pattern matching (word presence)")
    print(f"   Strategy 4: Category matching (account grouping)")
    
    filler = SemanticBalanceFiller(
        str(template_path),
        str(balance_path),
        inventory,
        fuzzy_threshold=0.50,
        enable_hybrid=True,  # NEW: Enable hybrid strategies
    )
    
    # Load data
    print(f"\n📥 Loading template and balance...")
    filler.load()
    print(f"✓ Template loaded: {len(filler.wb.sheetnames)} sheets")
    print(f"✓ Balance loaded: {len(filler.balance_accounts)} accounts")
    
    # Fill cells
    print(f"\n🔄 Filling cells (hybrid matching)...")
    try:
        output_path = Path('output/test_semantic_v6_output.xlsx')
        filled_count = filler.fill()
        print(f"✓ Filled {filled_count} cells")
    except Exception as e:
        print(f"\n✗ Error during filling: {e}")
        import traceback
        traceback.print_exc()
        raise
    
    # Validate
    print(f"\n📊 Validating results...")
    validator = SemanticFillerValidator(filler)
    
    # Save output
    print(f"\n💾 Saving output...")
    filler.save(output_path)
    print(f"✓ Output: {output_path}")
    
    # Generate reports
    print(f"\nGenerating validation reports...")
    validator.save_json_report(Path('output/reports/test_semantic_v6_report.json'))
    print(f"  ✓ JSON report: output/reports/test_semantic_v6_report.json")
    
    validator.save_html_report(Path('output/reports/test_semantic_v6_report.html'))
    print(f"  ✓ HTML report: output/reports/test_semantic_v6_report.html")
    
    # Console summary
    print(f"\n{'=' * 80}")
    print("📈 RESULTS SUMMARY (V6 HYBRID)")
    print(f"{'=' * 80}")
    
    summary = validator.generate_summary()
    print(summary)
    
    # Compare with v5
    print(f"\n{'=' * 80}")
    print("📊 V6 vs V5 COMPARISON")
    print(f"{'=' * 80}")
    print(f"\nV5 baseline:    572 cells (11.21%)")
    print(f"V6 target:      700+ cells (13-15%)")
    print(f"\nExpected improvement from hybrid strategies:")
    print(f"  - Hierarchical: +50-80 cells (prefix matching)")
    print(f"  - Pattern:      +20-40 cells (word presence)")
    print(f"  - Category:     +30-50 cells (account grouping)")
    print(f"  ────────────────────────")
    print(f"  Total target:   +100-170 cells → 672-742 total")
    print(f"  Success rate:   13.2-14.5%")


if __name__ == "__main__":
    main()
