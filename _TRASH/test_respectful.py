"""
Test script to verify respectful prefiller preserves template design
"""
from pathlib import Path
from datetime import datetime
from dsf_prefiller_respectful import DSFPrefillerRespectful
from dsf_general_info import DSF_InfosGenerales

print("\n" + "="*60)
print("DSF RESPECTFUL PREFILLER - DESIGN PRESERVATION TEST")
print("="*60 + "\n")

# Check if template exists
template = Path("DSF Normal standard.xlsx")
if not template.exists():
    print(f"❌ Template not found: {template}")
    exit(1)

print(f"✓ Template found: {template}")
print(f"  Size: {template.stat().st_size / 1024:.1f} KB")

# Initialize prefiller
print("\n📂 Initializing prefiller...")
prefiller = DSFPrefillerRespectful(
    template_path="DSF Normal standard.xlsx",
    mapping_path="dsf_prefill_mapping.json"
)

# Load template
print("   Loading template...")
prefiller.load_template()
print(f"   ✓ Sheets in template: {prefiller.wb.sheetnames}")

# Load mapping
print("   Loading mapping...")
prefiller.load_mapping()
print(f"   ✓ Sheets in mapping: {list(prefiller.mapping.keys())}")

# Create test data
print("\n📝 Creating test company data...")
company_data = DSF_InfosGenerales(
    denomination_sociale="GULFCAM S.A.S.",
    sigle_usuel="GULFCAM",
    num_identification_fiscale="R.C. 1998/SDE/CM",
    adresse_complete="Douala - Cameroun",
    exercice_debut=datetime(2024, 1, 1).date(),
    exercice_fin=datetime(2024, 12, 31).date(),
    date_arrete_comptes=datetime(2024, 12, 31).date(),
    systeme_comptable="Système Normal",
    forme_juridique="S.A.S.",
    registre_fiscal="SCIMPEX",
    centre_depot="Douala",
    ministere="MINFI",
    direction_generale="DGI"
)

# Fill data
print("\n📥 Filling data while preserving design...")
prefiller.prefill_entete(company_data)
prefiller.prefill_r1(company_data)
prefiller.prefill_r2(company_data)

# Save
print("\n💾 Saving output...")
output_file = prefiller.save("DSF_GULFCAM_RESPECTFUL.xlsx")

print("\n" + "="*60)
print("✅ TEST COMPLETE - Design should be preserved!")
print("="*60)
print(f"\nOutput file: {output_file}")
print(f"File size: {output_file.stat().st_size / 1024:.1f} KB")
print("\nNow OPEN THE FILE to verify:")
print("  ✓ ENTETE sheet - data filled correctly")
print("  ✓ All borders, colors, fonts preserved")
print("  ✓ Merged cells intact")
print("  ✓ No formatting destroyed\n")
