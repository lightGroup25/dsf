import json
from pathlib import Path

path = Path("data/dsf_inventory.json")
data = json.loads(path.read_text(encoding="utf-8"))

print("SHEETS LIST TO COPY PASTE:")
for sheet in data.get("sheets", {}).keys():
    # Print in Python Dict format for easy copy-paste
    print(f'    "{sheet}": "{sheet}",')
