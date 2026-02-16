import json
from pathlib import Path

path = Path("config/dsf_final_mapping.json")
data = json.loads(path.read_text(encoding="utf-8"))

level2 = data.get("LEVEL_2_INTERMEDIATE_TO_DSF", {})
print(f"Found {len(level2)} categories in Level 2:")
for cat in level2.keys():
    print(cat)
