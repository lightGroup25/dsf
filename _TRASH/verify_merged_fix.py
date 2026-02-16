"""
Test if data actually persisted in the merged cell fix output
"""
from openpyxl import load_workbook

output_file = "DSF_GULFCAM_MERGED_FIX.xlsx"

print("\n" + "="*70)
print("VERIFY DATA PERSISTENCE")
print("="*70)

# Load the output file
wb = load_workbook(output_file)

# Check ENTETE sheet
ws_entete = wb["ENTETE"]
print("\n📝 ENTETE Sheet:")
print(f"  A2 value: {ws_entete['A2'].value}")
print(f"  A2 value is None: {ws_entete['A2'].value is None}")
print(f"  A2 value length: {len(str(ws_entete['A2'].value)) if ws_entete['A2'].value else 0}")

# Check first few lines
if ws_entete['A2'].value:
    lines = str(ws_entete['A2'].value).split('\n')
    for i, line in enumerate(lines[:5]):
        print(f"    Line {i+1}: {line[:60]}...")

# Check INFORMATIONS GENERALES sheet
ws_info = wb["INFORMATIONS GENERALES"]
print("\n📋 INFORMATIONS GENERALES Sheet:")
for cell_ref in ["C2", "C3", "C4", "C5", "C6", "C7"]:
    value = ws_info[cell_ref].value
    print(f"  {cell_ref}: {value}")

print("\n" + "="*70)
if ws_entete['A2'].value and "GULFCAM" in str(ws_entete['A2'].value):
    print("✅ DATA PERSISTED SUCCESSFULLY!")
else:
    print("❌ DATA DID NOT PERSIST")
print("="*70 + "\n")
