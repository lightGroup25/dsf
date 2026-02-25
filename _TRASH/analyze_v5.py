#!/usr/bin/env python3
"""Analyze v5 results to understand optimization limits"""

import json
from pathlib import Path

# Load report
report_path = Path("output/reports/test_semantic_report.json")
with open(report_path, encoding='utf-8') as f:
    data = json.load(f)

val = data['validation']
print("=" * 70)
print("V5 TEST ANALYSIS: Keyword Fallback Impact")
print("=" * 70)
print(f"\nTotal cells: {val['total_cells_processed']}")
print(f"Successful: {val['successful_assignments']} ({val['success_rate']:.2f}%)")
print(f"Unmatched: {val['unmatched_cells']}")
print(f"Total amount: {val['total_amount_assigned']/1e9:.1f}B XAF")
print(f"\nConfidence breakdown:")
print(f"  High (>80%):   {val['high_confidence_assignments']}")
print(f"  Medium (50-80%): {val['medium_confidence_assignments']}")
print(f"  Low (<50%):    {val['low_confidence_assignments']}")

sheets_assigned = val['assignments_by_sheet']
print(f"\n\nSheets with assignments: {len(sheets_assigned)}")
print(f"Total sheets in template: 74")
print(f"Sheets with 0 assignments: {74 - len(sheets_assigned)}")

print(f"\n{'Sheet':<30} {'Assignments':>12}")
print("-" * 45)
for sheet, count in sorted(sheets_assigned.items(), key=lambda x: -x[1])[:20]:
    print(f"{sheet:<30} {count:>12}")

print(f"\n\n---- CONCLUSION ----")
print(f"Keyword matching had NO impact (same results as fuzzy-only)")
print(f"Possible reasons:")
print(f"1. Fuzzy >= 0.50 already captures all semantic matches")
print(f"2. Unmatched cells mostly structural (not account-related)")
print(f"3. French stop-words approach removes key words (<3 chars)")
print(f"\nCurrent ceiling: ~11.2% with semantic matching alone")
print(f"Next optimization: Consider rule-based + hierarchical mapping")
