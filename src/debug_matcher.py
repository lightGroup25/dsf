from balance_normalizer import BalanceNormalizer
from dsf_rule_config import load_rule_set
from dsf_rule_engine import RuleEngine
from pathlib import Path

# Paths
bal_path = Path("input/BALANCE AU 31 12 2024 VERSION DU 26 05 25.xlsx")
rules_path = Path("config/dsf_rule_generated.yaml")

import logging
logging.basicConfig(level=logging.INFO)

# Load
print("Loading Rules...")
rules = load_rule_set(rules_path)
print(f"Loaded {len(rules.rules)} rules.")

print("Loading Balance with Overrides...")
overrides = {
    "compte": 1,
    "label": 4,
    "debit_columns": [24, 21, 17, 11],
    "credit_columns": [27, 26, 20, 14, 13]
}
norm = BalanceNormalizer(bal_path, column_overrides=overrides, chunk_size=50)
print("Preparing parser...")
try:
    entries = list(norm.iterate())
except Exception as e:
    print(f"Error iterating balance: {e}")
    entries = []
print(f"Loaded {len(entries)} balance entries.")

# Match
print("Matching...")
matches = 0
ignored = []

for entry in entries:
    matched_rule = None
    for rule in rules.rules:
        if RuleEngine._row_matches_rule(entry, rule):
            matched_rule = rule
            break
    
    if matched_rule:
        val = RuleEngine._extract_value(entry, matched_rule)
        if val > 0:
            matches += 1
            # print(f"MATCH: {entry.compte} -> {matched_rule.id} (Val={val})")
        else:
            if matches < 5:
               pass
    else:
        ignored.append(entry)

print(f"Total Matches (Val > 0): {matches}")
print(f"Total Ignored: {len(ignored)}")
