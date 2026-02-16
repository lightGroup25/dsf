"""Tester et afficher le rapport du smart_general_filler pour R1/R2/R3"""
import sys
sys.path.insert(0, 'src')

from pathlib import Path
from smart_general_filler import SmartGeneralFiller
from dsf_general_info import DSF_InfosGenerales
from datetime import date

# Template à analyser
template = Path("templates/DSF Normal standard.xlsx")
if not template.exists():
    print(f"❌ Template introuvable: {template}")
    exit(1)

print(f"📄 Analyse du template: {template.name}\n")

# Créer le filler
filler = SmartGeneralFiller(template)
filler.load()

print(f"✅ {len(filler.input_zones)} zones de saisie détectées\n")

# Filtrer pour R1, R2, R3
r_zones = [z for z in filler.input_zones if any(s in z.sheet for s in ["R1", "R2", "R3"])]

print(f"📋 Zones dans R1/R2/R3: {len(r_zones)}\n")
print("=" * 100)

# Grouper par feuille
by_sheet = {}
for z in r_zones:
    if z.sheet not in by_sheet:
        by_sheet[z.sheet] = []
    by_sheet[z.sheet].append(z)

# Afficher par feuille
for sheet in sorted(by_sheet.keys()):
    zones = by_sheet[sheet]
    print(f"\n📑 {sheet} - {len(zones)} zones détectées")
    print("-" * 100)
    
    for z in zones[:15]:  # Max 15 zones par feuille
        print(f"  {z.cell_ref:6} | Type: {z.zone_type:10} | Label: {z.label[:50]:50} | Context: {z.context[:30]}")

# Créer des infos de test
info = DSF_InfosGenerales(
    raison_sociale="ENTREPRISE TEST SA",
    sigle="ETEST",
    adresse="123 RUE PRINCIPALE",
    ville="DOUALA",
    pays="CAMEROUN",
    num_contribuable="M0123456789Z",
    exercice_debut=date(2024, 1, 1),
    exercice_fin=date(2024, 12, 31),
    duree_mois=12
)

# Remplir et obtenir le rapport
print(f"\n{'=' * 100}")
print("🔄 Test de remplissage avec données fictives")
print("=" * 100)

filled = filler.fill(info)
print(f"\n✅ {filled} zones remplies")

# Afficher le rapport détaillé
report = filler.get_report()

print(f"\n📊 RAPPORT PAR FEUILLE:")
print("=" * 100)

for sheet, data in report["by_sheet"].items():
    if any(s in sheet for s in ["R1", "R2", "R3"]):
        print(f"\n📑 {sheet}")
        print(f"   Détectées: {data['detected']} | Remplies: {data['filled']}")
        
        # Afficher les zones remplies
        filled_zones = [z for z in data["zones"] if z["matched_field"]]
        if filled_zones:
            print(f"\n   Zones REMPLIES:")
            for z in filled_zones[:10]:
                print(f"      {z['cell']:6} ← {z['matched_field']:25} (conf={z['confidence']:.2f}) | Label: {z['label'][:40]}")

filler.close()
