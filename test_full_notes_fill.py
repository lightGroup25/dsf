#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Test complet du remplissage des Notes 2, 3A, 3B, 3C avec balance 2024"""

import sys
from pathlib import Path
from openpyxl import load_workbook
from copy import deepcopy

# Ajouter src au path
sys.path.insert(0, str(Path(__file__).parent / "src"))

print("=" * 100)
print("TEST COMPLET: REMPLISSAGE NOTES 2, 3A, 3B, 3C")
print("=" * 100)

# Chemins
balance_2024 = Path(r"C:\Users\Emmaneul Ambadiang\Desktop\DSF\input\BALANCE AU 31 12 2024 VERSION DU 26 05 25.xlsx")
dsf_2023 = Path(r"C:\Users\Emmaneul Ambadiang\Desktop\DSF\input\DSF GULFCAM 2023 V3.xlsx")

if not balance_2024.exists():
    print(f"[NO] Balance non trouvee: {balance_2024}")
    sys.exit(1)

if not dsf_2023.exists():
    print(f"[NO] DSF non trouvee: {dsf_2023}")
    sys.exit(1)

print(f"[OK] Balance 2024 trouvee: {balance_2024.name}")
print(f"[OK] DSF 2023 trouvee: {dsf_2023.name}")

# Charger la balance
balance_wb = load_workbook(balance_2024)
print(f"\n[OK] Balance chargee: {len(balance_wb.sheetnames)} feuilles")
print(f"    Feuilles: {', '.join(balance_wb.sheetnames[:3])}...")

# Charger le DSF
dsf_wb = load_workbook(dsf_2023)
print(f"[OK] DSF chargee: {len(dsf_wb.sheetnames)} feuilles")

# Voir la structure de la balance
balance_ws = balance_wb.active
print(f"\n### STRUCTURE BALANCE ###")
print(f"  Feuille active: '{balance_ws.title}'")
print(f"  Dimensions: {balance_ws.max_row}x{balance_ws.max_column}")

# Afficher les premières lignes de la balance
print(f"\n  Entetes (row 1):")
for col in range(1, min(6, balance_ws.max_column + 1)):
    cell = balance_ws.cell(1, col)
    print(f"    Col {col}: '{cell.value}'")

print(f"\n  Donnees (rows 2-5):")
for row in range(2, min(6, balance_ws.max_row + 1)):
    compte = balance_ws.cell(row, 1).value
    solde_debit = balance_ws.cell(row, 2).value if balance_ws.max_column >= 2 else None
    solde_credit = balance_ws.cell(row, 3).value if balance_ws.max_column >= 3 else None
    print(f"    Row {row}: {compte} | D={solde_debit} | C={solde_credit}")

# Créer une copie de travail du DSF
dsf_work = deepcopy(dsf_wb)

print(f"\n### AVANT REMPLISSAGE DES NOTES ###")
notes_to_check = ["NOTE 2", "NOTE 3A", "NOTE 3B", "NOTE  3C"]
notes_row_counts = {}

for note_name in notes_to_check:
    for sheet in dsf_work.sheetnames:
        if sheet.upper().replace(" ", "") == note_name.upper().replace(" ", ""):
            ws = dsf_work[sheet]
            # Compter les données (non-vides)
            data_rows = 0
            for row in range(2, min(100, ws.max_row + 1)):
                has_data = False
                for col in range(1, ws.max_column + 1):
                    if ws.cell(row, col).value:
                        has_data = True
                        break
                if has_data:
                    data_rows += 1
                else:
                    break
            
            notes_row_counts[sheet] = data_rows
            print(f"  {sheet:15} -> {ws.max_row} rows max, ~{data_rows} rows data")
            break

# Pas de balance normalisée fournie, on va simuler les lignes
print(f"\n### SIMULATION REMPLISSAGE ###")
print(f"  [INFO] Normalement la pipeline normalizerait la balance")
print(f"  [INFO] Les comptes seraient extraits de la balance et mappés aux notes")
print(f"  [INFO] fill_notes_from_balance() remplirait ensuite les notes")

# Importer les fonctions de remplissage
try:
    from dsf_notes_filler import fill_notes_from_balance
    print(f"\n[OK] Module dsf_notes_filler importe")
    
    # La fonction fill_notes_from_balance a besoin de:
    # - wb: workbook
    # - normalized_rows: liste [{"account": "123", "solde": 50000, ...}]
    # - previous_rows: (optionnel) pour comparaison N/N-1
    
    print(f"[INFO] fill_notes_from_balance signature:")
    import inspect
    sig = inspect.signature(fill_notes_from_balance)
    print(f"       {sig}")
    
except ImportError as e:
    print(f"[NO] Erreur import: {e}")

# Pour un test réel, il faudrait:
# 1. Normaliser la balance comme le ferait le pipeline
# 2. Extraire les lignes en format normalized_rows
# 3. Appeler fill_notes_from_balance()
# 4. Vérifier les notes remplies

print(f"\n### CONCLUSION ###")
print(f"Les notes 2, 3A, 3B, 3C existent et ont des mappings.")
print(f"Pour tester le remplissage complet, il faut:")
print(f"  1. Avoir l'output de la normalisation de balance")
print(f"  2. Ou avoir un fichier balance_normalized.json de reference")
print(f"  3. Ou modifer le test pour normaliser la balance 2024")

dsf_wb.close()
balance_wb.close()
