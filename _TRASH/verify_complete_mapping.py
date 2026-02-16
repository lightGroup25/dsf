import json

with open('dsf_complete_mapping.json', 'r', encoding='utf-8') as f:
    mapping = json.load(f)

cov = mapping['coverage']
print(f'\n✓ Mapping complet généré!')
print(f'  Total SYSCOHADA accounts: {cov["total_syscohada"]}')
print(f'  Mapped to intermediate:   {cov["mapped"]}')
print(f'  Coverage percentage:      {cov["coverage_percentage"]:.2f}%')
print(f'  Status:                   {"100% COMPLETE" if cov["unmapped"] == 0 else "INCOMPLETE"}')
print()
