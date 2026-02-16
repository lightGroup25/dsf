#!/usr/bin/env python3
"""
Quick V6 test - Compare results with V5
"""

import sys
from pathlib import Path
from decimal import Decimal

sys.path.insert(0, str(Path('src')))

from semantic_balance_filler import SemanticBalanceFiller
from semantic_validators import SemanticFillerValidator
from dsf_inventory import DSFInventory

def main():
    print("=" * 80)
    print("🧪 TEST V6 QUICK - HYBRID OPTIMIZATION")
    print("=" * 80)
    
    template_path = Path('templates/DSF Normal standard.xlsx')
    balance_path = Path('input/BALANCE AU 31 12 2024 VERSION DU 26 05 25.xlsx')
    inventory_path = Path('data/dsf_inventory.json')
    output_path = Path('output/test_semantic_v6_output.xlsx')
    report_path = Path('output/reports/test_semantic_v6_report.json')
    
    try:
        # Load
        print(f"\n📖 Loading...")
        inventory = DSFInventory.from_json(inventory_path)
        
        # Create filler
        print(f"🔧 Initializing filler (hybrid=True)...")
        filler = SemanticBalanceFiller(
            str(template_path),
            str(balance_path),
            inventory,
            fuzzy_threshold=0.50,
            enable_hybrid=True,
        )
        
        # Load and fill
        print(f"📥 Loading and filling...")
        filler.load()
        filled_count = filler.fill()
        print(f"✓ Filled: {filled_count} cells")
        
        # Save
        print(f"💾 Saving...")
        filler.save(output_path)
        
        # Validate
        print(f"\n📊 Validating...")
        validator = SemanticFillerValidator(filler)
        report = validator.validate()
        
        # Results
        print(f"\n{'=' * 80}")
        print("📈 V6 RESULTS")
        print(f"{'=' * 80}")
        print(f"\nTotal cells:        {report.total_cells_processed}")
        print(f"Successful:         {report.successful_assignments} ({report.success_rate:.2f}%)")
        print(f"Unmatched:          {report.unmatched_cells}")
        print(f"Total amount:       {report.total_amount_assigned/1e9:.1f}B XAF")
        print(f"High confidence:    {report.high_confidence_assignments}")
        print(f"Medium confidence:  {report.medium_confidence_assignments}")
        
        # Compare with V5
        print(f"\n{'=' * 80}")
        print("📊 V6 vs V5 COMPARISON")
        print(f"{'=' * 80}")
        print(f"V5: 572 cells (11.21%)")
        print(f"V6: {report.successful_assignments} cells ({report.success_rate:.2f}%)")
        
        improvement = report.successful_assignments - 572
        improvement_pct = report.success_rate - 11.21
        
        if improvement > 0:
            print(f"\n✅ IMPROVEMENT: +{improvement} cells (+{improvement_pct:.2f}%)")
        elif improvement == 0:
            print(f"\n⚠ NO CHANGE: Hybrid strategies did not help")
        else:
            print(f"\n❌ DEGRADATION: -{abs(improvement)} cells")
        
        # Save report
        print(f"\n📄 Saving report...")
        validator.save_json_report(report_path)
        print(f"✓ Report: {report_path}")
        
    except Exception as e:
        print(f"\n❌ ERROR: {e}")
        import traceback
        traceback.print_exc()
        return 1
    
    return 0

if __name__ == "__main__":
    sys.exit(main())
