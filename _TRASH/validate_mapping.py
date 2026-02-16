import json

with open('dsf_final_mapping.json') as f:
    data = json.load(f)

print("✅ Valid JSON")
print(f"Keys: {list(data.keys())}")
print(f"Total Level 1: {len(data['LEVEL_1_SYSCOHADA_TO_INTERMEDIATE'])}")
print(f"Sample entries:")
for i, (key, val) in enumerate(list(data['LEVEL_1_SYSCOHADA_TO_INTERMEDIATE'].items())[:3]):
    print(f"  {key}: {val}")
