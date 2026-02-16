#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import sys
from pathlib import Path
sys.path.insert(0, str(Path('src')))

from semantic_balance_filler import SemanticBalanceFiller
from dsf_inventory import DSFInventory

inventory = DSFInventory.from_json(Path('data/dsf_inventory.json'))
filler = SemanticBalanceFiller(
    'templates/DSF Normal standard.xlsx',
    'input/BALANCE AU 31 12 2024 VERSION DU 26 05 25.xlsx',
    inventory,
    fuzzy_threshold=0.50
)
filler.load()
print(f'Fuzzy threshold: {filler.fuzzy_threshold}')

# Test scores
test_labels = [
    ('Stocks', 'STOCKS FINIS'),
    ('Stocks', 'RÉSERVES'),
    ('Capital', 'CAPITAL SOCIAL'),
    ('Interest', 'INTERESSEMENT'),
]

for label, account_label in test_labels:
    score = filler._fuzzy_similarity(label.lower(), account_label.lower())
    matches_at_50 = "✓" if score >= 0.50 else "✗"
    matches_at_70 = "✓" if score >= 0.70 else "✗"
    print(f'{label:15s} vs {account_label:25s}: {score:.2f} (0.50: {matches_at_50}, 0.70: {matches_at_70})')
