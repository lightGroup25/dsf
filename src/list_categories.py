import json
from pathlib import Path

path = Path("config/dsf_final_mapping.json")
data = json.loads(path.read_text(encoding="utf-8"))

cats = set()
# The structure is LEVEL_1... -> key -> intermediate_category
l1 = data.get("LEVEL_1_SYSCOHADA_TO_INTERMEDIATE", {})
for k, v in l1.items():
    cats.add(v["intermediate_category"])

print('\n'.join(sorted(list(cats))))

