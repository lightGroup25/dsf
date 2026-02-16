# -*- coding: utf-8 -*-
"""Test du remplissage des tableaux dirigeants et conseil dans la fiche R3"""
import sys
sys.path.insert(0, 'src')

from pathlib import Path
from smart_general_filler import SmartGeneralFiller
from dsf_general_info import DSF_InfosGenerales, Dirigeant, Conseil_Administration
from datetime import date

print("=" * 80)
print("TEST: Remplissage des tableaux R3 (Dirigeants + Conseil)")
print("=" * 80)

# Template
template = Path("templates/DSF Normal standard.xlsx")
output = Path("output/TEST_R3_TABLES.xlsx")

if not template.exists():
    print(f"Erreur: Template introuvable: {template}")
    exit(1)

# Créer des données de test avec dirigeants et conseil
info = DSF_InfosGenerales(
    denomination_sociale="ENTREPRISE TEST SA",
    sigle_usuel="ETEST",
    adresse_complete="123 RUE PRINCIPALE, DOUALA",
    num_identification_fiscale="M0123456789Z",
    exercice_debut=date(2024, 1, 1),
    exercice_fin=date(2024, 12, 31),
    date_arrete_comptes=date(2025, 1, 31),
    duree_mois=12,
    pays="CAMEROUN",
    ville="DOUALA",
    
    # DIRIGEANTS (3 personnes)
    dirigeants=[
        Dirigeant(
            nom="DUPONT",
            prenoms="Jean-Pierre",
            qualite="PRESIDENT DIRECTEUR GENERAL",
            num_identification_fiscale="M050900027774W",
            adresse="BP 1234, DOUALA, CAMEROUN"
        ),
        Dirigeant(
            nom="MARTIN",
            prenoms="Marie-Claire",
            qualite="DIRECTEUR GENERAL ADJOINT",
            num_identification_fiscale="M050900027775W",
            adresse="BP 5678, DOUALA, CAMEROUN"
        ),
        Dirigeant(
            nom="BERNARD",
            prenoms="Luc",
            qualite="DIRECTEUR FINANCIER",
            num_identification_fiscale="M050900027776W",
            adresse="BP 9012, DOUALA, CAMEROUN"
        ),
    ],
    
    # CONSEIL D'ADMINISTRATION (4 membres)
    conseil_administration=[
        Conseil_Administration(
            nom="LEFEBVRE",
            prenoms="Philippe",
            qualite="PRESIDENT DU CONSEIL",
            adresse="BP 2000, YAOUNDE, CAMEROUN"
        ),
        Conseil_Administration(
            nom="DURAND",
            prenoms="Sophie",
            qualite="MEMBRE DU CONSEIL",
            adresse="BP 3000, DOUALA, CAMEROUN"
        ),
        Conseil_Administration(
            nom="PETIT",
            prenoms="Michel",
            qualite="MEMBRE DU CONSEIL",
            adresse="BP 4000, DOUALA, CAMEROUN"
        ),
        Conseil_Administration(
            nom="ROUX",
            prenoms="Catherine",
            qualite="MEMBRE DU CONSEIL",
            adresse="BP 5000, YAOUNDE, CAMEROUN"
        ),
    ]
)

print(f"\nDonnees de test:")
print(f"  Dirigeants: {len(info.dirigeants)}")
for d in info.dirigeants:
    print(f"    - {d.nom} {d.prenoms} ({d.qualite})")

print(f"\n  Conseil d'administration: {len(info.conseil_administration)}")  
for c in info.conseil_administration:
    print(f"    - {c.nom} {c.prenoms} ({c.qualite})")

# Remplir
print(f"\nRemplissage...")
filler = SmartGeneralFiller(template)
filler.load()

filled = filler.fill(info)
print(f"\nCellules remplies: {filled}")

# Sauvegarder
output.parent.mkdir(parents=True, exist_ok=True)
filler.save(output)
filler.close()

print(f"\nFichier genere: {output}")
print(f"\nVerification:")
print(f"  1. Ouvrez: {output}")
print(f"  2. Allez sur 'Fiche R3'")
print(f"  3. Verifiez lignes 10-12 (dirigeants) et lignes 33-36 (conseil)")
print(f"\n{'=' * 80}")
print("Test termine!")
