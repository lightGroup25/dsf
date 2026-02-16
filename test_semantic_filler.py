#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Test script for Semantic Balance Filler
Validates the new approach on a subset of data
"""
import sys
import io
from pathlib import Path

# Fix encoding for Windows console (emojis)
if sys.stdout.encoding != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

# Add src to path
src_path = Path(__file__).parent / "src"
sys.path.insert(0, str(src_path))

from dsf_inventory import DSFInventory
from semantic_balance_filler import SemanticBalanceFiller
from semantic_validators import SemanticFillerValidator

def main():
    print("=" * 70)
    print("🧪 TEST - SEMANTIC BALANCE FILLER")
    print("=" * 70)

    # Configuration
    template_path = Path("templates/DSF Normal standard.xlsx")
    balance_path = Path("input/BALANCE AU 31 12 2024 VERSION DU 26 05 25.xlsx")
    inventory_path = Path("data/dsf_inventory.json")
    output_path = Path("output/test_semantic_output.xlsx")
    report_json = Path("output/reports/test_semantic_report.json")
    report_html = Path("output/reports/test_semantic_report.html")

    # Check inputs
    if not template_path.exists():
        print(f"❌ Template not found: {template_path}")
        return
    if not balance_path.exists():
        print(f"❌ Balance not found: {balance_path}")
        return
    if not inventory_path.exists():
        print(f"❌ Inventory not found: {inventory_path}")
        return

    print(f"✓ Template: {template_path}")
    print(f"✓ Balance: {balance_path}")
    print(f"✓ Inventory: {inventory_path}")

    # Load inventory
    print("\n📖 Loading inventory...")
    inventory = DSFInventory.from_json(inventory_path)
    total_fields = sum(len(fields) for fields in inventory._fields_by_sheet.values())
    print(f"✓ Inventory loaded: {len(inventory._fields_by_sheet)} sheets, {total_fields} fields")

    # Create filler
    print("\n🔧 Initializing Semantic Balance Filler...")
    filler = SemanticBalanceFiller(
        template_path,
        balance_path,
        inventory,
        fuzzy_threshold=0.50,  # Reduced from 0.70 for better matching (was too strict)
    )

    # Load data
    print("📥 Loading template and balance...")
    filler.load()
    print(f"✓ Template loaded: {len(filler.wb.sheetnames)} sheets")
    print(f"✓ Balance loaded: {len(filler.balance_accounts)} accounts")

    # Fill
    print("\n🔄 Filling cells (semantic matching)...")
    try:
        filled_count = filler.fill()
        print(f"✓ Filled {filled_count} cells")
    except Exception as e:
        print(f"❌ ERROR during fill: {type(e).__name__}: {e}")
        import traceback
        traceback.print_exc()
        return

    # Validate
    print("\n✅ Validating results...")
    validator = SemanticFillerValidator(filler)
    report = validator.validate()

    print(f"✓ Total cells: {report.total_cells_processed}")
    print(f"✓ Successful: {report.successful_assignments}")
    print(f"✓ Unmatched: {report.unmatched_cells}")
    print(f"✓ Success rate: {report.success_rate:.1f}%")
    print(f"✓ Total amount: {report.total_amount_assigned:,.0f}")
    print(f"✓ High confidence: {report.high_confidence_assignments}")
    print(f"✓ Medium confidence: {report.medium_confidence_assignments}")
    print(f"✓ Low confidence: {report.low_confidence_assignments}")

    # Generate reports
    print("\n📊 Generating reports...")
    validator.generate_json_report(report_json)
    validator.generate_html_report(report_html)
    print(f"✓ JSON report: {report_json}")
    print(f"✓ HTML report: {report_html}")

    # Save output
    print("\n💾 Saving output...")
    filler.save(output_path)
    print(f"✓ Output: {output_path}")

    # Print console report
    print("\n" + "=" * 70)
    validator.print_console_report()
    print("=" * 70)

if __name__ == "__main__":
    main()
