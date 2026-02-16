#!/usr/bin/env python3
"""
PROPER COLUMN MAPPING STRATEGY
Map DSF column headers to Balance column sources
Instead of fuzzy matching everything blindly
"""

print("=" * 80)
print("DSF COLUMN MAPPING STRATEGY")
print("=" * 80)

# Based on real analysis:
# DSF Columns (NOTE 3A etc):
#   Col 4: "MONTANT BRUT A L'OUVERTURE" 
#   Col 5: "ACQUISITIONS APPORTS CREATIONS"
#   Col 6-9: Various movements
#   Col 10: "MONTANT BRUT A LA CLOTURE"

# Balance Columns:
#   Col 1: Account number
#   Col 4: Account label
#   Col 10: Mouvements 31/12/23 - Débit
#   Col 13: Mouvements 31/12/23 - Crédit
#   Col 14: Soldes final - Débit (TARGET for "MONTANT BRUT A CLOTURE")
#   Col 27: Soldes final - Crédit

# MAPPING STRATEGY:
dsf_column_mapping = {
    "MONTANT BRUTE A L'OUVERTURE": {
        "balance_sources": [12, 13],  # Opening balance debit/credit (columns 12-13?)
        "type": "opening_balance",
        "debit_col": None,  # Need to find
        "credit_col": None,
    },
    "ACQUISITIONS APPORTS CREATIONS": {
        "balance_sources": [10, 20],  # Mouvements (movements)
        "type": "acquisition_movement",
        "debit_col": None,
        "credit_col": None,
    },
    "VIREMENTS DE POSTE A POSTE": {
        "balance_sources": [],  # Transfer movements
        "type": "transfer_movement",
    },
    "SUITE A UNE REEVALUATION PRATIQUEE AU COURS DE L'EXERCICE": {
        "balance_sources": [],  # Revaluation
        "type": "revaluation",
    },
    "CESSIONS SCISSIONSHORS SERVICE": {
        "balance_sources": [],  # Asset disposals/removals
        "type": "disposal",
    },
    "MONTANT BRUT A LA CLOTURE": {
        "balance_sources": [14, 27],  # FINAL BALANCE Debit/Credit columns
        "type": "closing_balance",
        "debit_col": 14,  # Column with debit balances
        "credit_col": 27,  # Column with credit balances
    }
}

print("\nDSF Column Headers → Balance Column Sources:")
print("-" * 80)

for header, mapping in dsf_column_mapping.items():
    print(f"\n{header}")
    print(f"  Type: {mapping['type']}")
    print(f"  Balance sources: {mapping['balance_sources']}")
    if 'debit_col' in mapping and mapping['debit_col']:
        print(f"  Debit column: {mapping['debit_col']}")
        print(f"  Credit column: {mapping['credit_col']}")

print("\n" + "=" * 80)
print("HOW THIS SHOULD WORK:")
print("=" * 80)
print("""
OLD (WRONG):
  1. See cell in NOTE 3A
  2. Match row label "STOCKS" to balance row
  3. Take first non-zero amount
  4. Fill cell

NEW (CORRECT):
  1. See cell in NOTE 3A, column 10 ("MONTANT BRUT A LA CLOTURE")
  2. Understand: I need FINAL BALANCE of this account
  3. Match row label "STOCKS" to balance account
  4. Look in balance COLUMN 14 or 27 (closing debit/credit)
  5. Take amount from correct column based on account natural side
  6. Fill cell

ADVANTAGE:
  - No more "same value in all columns"
  - Column headers GUIDE the mapping
  - We use SPECIFIC columns, not guessing
  - Business logic: "Closing balance" ≠ "Acquisitions"
""")
