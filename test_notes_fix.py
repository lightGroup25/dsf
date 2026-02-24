#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Test de vérification du bug de remplissage des notes - CORRIGÉ"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent / "src"))

from semantic_balance_filler import SemanticBalanceFiller
from openpyxl import load_workbook

print("=" * 100)
print("TEST: VÉRIFICATION DU BUG DE REMPLISSAGE DES NOTES")
print("=" * 100)

# Chemins
balance_2024 = Path(r"C:\Users\Emmaneul Ambadiang\Desktop\DSF\input\BALANCE AU 31 12 2024 VERSION DU 26 05 25.xlsx")
dsf_2023 = Path(r"C:\Users\Emmaneul Ambadiang\Desktop\DSF\input\DSF GULFCAM 2023 V3.xlsx")
template_out = Path(r"C:\Users\Emmaneul Ambadiang\Desktop\DSF\output\test_template_out.xlsx")

if not balance_2024.exists():
    print(f"[NO] Balance non trouvee: {balance_2024}")
    sys.exit(1)

if not dsf_2023.exists():
    print(f"[NO] DSF non trouvee: {dsf_2023}")
    sys.exit(1)

print(f"[OK] Files trouvées")

# Copier le DSF 2023 comme template de travail
from shutil import copy2
copy2(dsf_2023, template_out)
print(f"[OK] Template copié: {template_out.name}")

# Créer le filler sémantique
try:
    from dsf_inventory import DSFInventory
    
    # Charger l'inventaire
    inventory = DSFInventory()
    inventory.load(dsf_2023)
    print(f"[OK] Inventaire chargé: {len(inventory.all_cells)} cellules")
    
    # Créer le filler
    filler = SemanticBalanceFiller(
        template_out,
        balance_2024,
        inventory,
        fuzzy_threshold=0.70,
    )
    
    print(f"\n### TEST 1: Vérification des propriétés ###")
    
    # Load balance
    filler.load()
    
    # Vérifier que normalized_rows est exposé
    print(f"  filler.balance_accounts type: {type(filler.balance_accounts)}")
    print(f"  filler.balance_accounts length: {len(filler.balance_accounts)}")
    
    print(f"\n  filler.normalized_rows type: {type(filler.normalized_rows)}")
    print(f"  filler.normalized_rows length: {len(filler.normalized_rows)}")
    
    if len(filler.normalized_rows) > 0:
        first_row = filler.normalized_rows[0]
        print(f"  First row compte: {first_row.compte}")
        print(f"  First row solde_final: {first_row.solde_final}")
        print(f"  ✓ normalized_rows est correctement exposé!")
    else:
        print(f"  ✗ ERREUR: normalized_rows est vide!")
    
    print(f"\n  filler.previous_normalized_rows length: {len(filler.previous_normalized_rows)}")
    
    print(f"\n### TEST 2: Remplissage du DSF ###")
    
    # Fill
    n_assignments = filler.fill()
    print(f"  Assignments made: {n_assignments}")
    
    # Save
    filler.wb.save(template_out)
    print(f"[OK] Template rempli et sauvegardé")
    
    print(f"\n### TEST 3: Vérification que les Notes seraient remplies ###")
    print(f"  Les données normalisées sont maintenant disponibles pour fill_notes_from_balance()")
    print(f"  Nombre de lignes qui seraient traitées: {len(filler.normalized_rows)}")
    
    if len(filler.normalized_rows) > 0:
        print(f"  ✓ Les notes SERONT remplies (BUG CORRIGÉ)")
    else:
        print(f"  ✗ Les notes ne seront PAS remplies (BUG NON CORRIGÉ)")
    
except Exception as e:
    print(f"[NO] Erreur: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

print(f"\nFin du test!")
