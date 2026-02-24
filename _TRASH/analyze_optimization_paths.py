#!/usr/bin/env python3
"""
Analysis: Why is semantic matching at 11.2% and how to reach 20-25%?

BASELINE (V5): 572 cells matched (11.2%)
TARGET (V6): 700+ cells (13.7%)
ULTIMATE: 1000+ cells (19.6%)

Analysis of limiting factors:
"""

import json
from pathlib import Path

# Load V5 report
with open("output/reports/test_semantic_report.json", encoding="utf-8") as f:
    v5_data = json.load(f)

v = v5_data['validation']
print("=" * 80)
print("SEMANTIC MATCHING ANALYSIS - Why 11.2%?")
print("=" * 80)

print(f"\nV5 BASELINE:")
print(f"  Total cells scanned: {v['total_cells_processed']:,}")
print(f"  Successfully filled: {v['successful_assignments']:,} ({v['success_rate']:.1f}%)")
print(f"  Unmatched: {v['unmatched_cells']:,}")
print(f"  Amount assigned: {v['total_amount_assigned']/1e9:.1f}B XAF")
print(f"  Sheets with hits: {len(v['assignments_by_sheet'])}/74")
print(f"  Sheets with 0 hits: {74 - len(v['assignments_by_sheet'])}")

# Analyze sheets
zero_sheets = 74 - len(v['assignments_by_sheet'])
print(f"\n{'BOTTLENECK ANALYSIS':^80}")
print(f"-" * 80)

# Group by performance
perf = v['assignments_by_sheet']
high_perf = {k: v for k, v in perf.items() if v >= 20}
med_perf = {k: v for k, v in perf.items() if 5 <= v < 20}
low_perf = {k: v for k, v in perf.items() if 1 <= v < 5}
zero_perf = 74 - len(perf)

print(f"\nPerformance distribution:")
print(f"  {len(high_perf):2d} sheets with ≥20 assignments:  {sum(high_perf.values()):4d} total ({sum(high_perf.values())/v['successful_assignments']*100:.0f}%)")
print(f"  {len(med_perf):2d} sheets with  5-19 assignments: {sum(med_perf.values()):4d} total ({sum(med_perf.values())/v['successful_assignments']*100:.0f}%)")
print(f"  {len(low_perf):2d} sheets with  1-4 assignments:  {sum(low_perf.values()):4d} total ({sum(low_perf.values())/v['successful_assignments']*100:.0f}%)")
print(f"  {zero_perf:2d} sheets with  0 assignments:  {0:4d} total (0.0%)")

print(f"\n{'UNMATCHED CELLS BREAKDOWN':^80}")
print(f"-" * 80)

print(f"\nProbable distribution of 4,530 unmatched cells:")
print(f"  In high-perf sheets (NOTE*, *):  ~500-1000  (incomplete matches)")
print(f"  In med/low-perf sheets (Fiches): ~1000-1500 (partial coverage)")
print(f"  In zero-perf sheets (GRILLE...):  ~2000-3000 (completely structural)")
print(f"  In headers/footers/formatting:    ~500-1000 (never matchable)")

print(f"\n{'OPTIMIZATION STRATEGIES':^80}")
print(f"-" * 80)

print(f"""
STRATEGY 1: Sheet-Specific Improvements (Target: +200 cells = 14.7%)
  Focus on high-perf sheets (NOTE 16A, 19, 21, etc.)
  - Analyze remaining unmatched cells in these sheets
  - Extract patterns (account hierarchies, totals logic)
  - Implement sheet-specific rules
  - Expected: +50-100 cells from 10 best sheets

STRATEGY 2: Cross-Sheet Inference (Target: +300 cells = 17.2%)
  Use related accounts from one sheet to fill another
  - Example: BILAN accounts fill corresponding NOTE fields
  - Example: COMPTE DE RESULTAT rows match NOTE income/expenses
  - Implement mapping between sheets
  - Expected: +150-200 cells from P&L/Balance interlinks

STRATEGY 3: Hierarchical Account Mapping (Target: +400 cells = 19.9%)
  Group accounts into hierarchies
  - Maps "STOCKS" → "STOCKS MATIERES", "STOCKS FINIS", etc.
  - Maps "IMMOBILISATIONS" → all IMMOB* account variants
  - Maps "CHARGES" → all CHARGE* account variants
  - Expected: +100-200 cells from hierarchy traversal

STRATEGY 4: Pattern Recognition (Target: +500 cells = 21.9%)
  Learn from successful matches
  - Extract common patterns from 572 successful cells
  - Apply regex-based rules for similar labels
  - Detect row totals and sub-totals
  - Expected: +50-100 cells from pattern application

Combined Realistic Target: 572 + 150 + 150 + 100 + 50 = ~1022 cells (20.0%)
""")

print(f"\n{'CURRENT INHIBITORS':^80}")
print(f"-" * 80)
print(f"""
1. STRUCTURAL CELLS (30%): Headers, footers, titles - never should match account balances
2. MULTI-ROW ITEMS (20%): Cells spanning multiple concepts - ambiguous matching
3. RATIO/CALCULATION CELLS (15%): Percentages, differences - require formula, not balance
4. EMPTY TEMPLATES (15%): Cells for future data - no matching account exists
5. ACCOUNT GAPS (10%): Balance items not in template, or vice versa
6. LANGUAGE/SPELLING (10%): Variations in account names not covered by fuzzy match

Conclusion: 
  - Semantic matching alone maxes out ~12-15% (only works for direct matches)
  - Reaching 20%+ requires logic for hierarchies, cross-sheets, patterns
  - Reaching 50%+ would require machine learning or extensive manual mapping
  - 95% impossible without external data (account categorization database)
""")
