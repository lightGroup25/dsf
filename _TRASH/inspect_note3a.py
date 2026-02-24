import json
from pathlib import Path

path = Path("data/dsf_inventory.json")
data = json.loads(path.read_text(encoding="utf-8"))

note_3a = data["sheets"].get("NOTE 3A", {})
print("All Columns in NOTE 3A:")
for c in note_3a.get("input_columns", []):
    print(f"Col {c['letter']} ({c['column']}): {c['header_value'].replace(chr(10), ' ')}")
print("-" * 50)
print("First 10 Fields:")
for field in note_3a.get("fields", [])[:10]:
    print(f"Row {field['row']}, Col {field['column_letter']} ({field['cell']}) - Label: '{field.get('label')}'")
