#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Lire et analyser TOUS les libellés de colonnes du DSF
Pour voir les patterns et comprendre la logique
"""
import openpyxl
from openpyxl.utils import get_column_letter

dsf_template = r"c:\Users\Emmaneul Ambadiang\Desktop\DSF\templates\DSF Normal standard.xlsx"

wb = openpyxl.load_workbook(dsf_template, data_only=False)

# Lister les feuilles pertinentes (notes d'actif/passif)
relevant_sheets = [
    "NOTE 3A",  # Immobilisations brutes
    "NOTE 3B",  # Dépréciations/amortissements
    "NOTE 4 ",  # Stocks
    "NOTE 5 ",  # Créances
    "NOTE 6 ",  # Trésorerie
]

print("=" * 150)
print("ANALYSE DES LIBELLES DE COLONNES DSF")
print("=" * 150)

for sheet_name in relevant_sheets:
    if sheet_name not in wb.sheetnames:
        # Essayer sans espace final
        sheet_name_alt = sheet_name.rstrip()
        if sheet_name_alt in wb.sheetnames:
            sheet_name = sheet_name_alt
        else:
            print(f"\n[SKIPPED] {sheet_name} not found")
            continue
    
    ws = wb[sheet_name]
    
    print(f"\n{'=' * 150}")
    print(f"FEUILLE: {sheet_name}")
    print(f"{'=' * 150}")
    
    # Lire les 3 premières lignes (titres généralement en rows 8-10)
    print("\nRanges de titres (rows 8-10):")
    print("-" * 150)
    
    header_data = {}
    for row_num in [8, 9, 10]:
        print(f"\nRow {row_num}:")
        row_data = []
        for col_num in range(1, 16):
            cell = ws.cell(row_num, col_num)
            val = cell.value
            if val:
                col_letter = get_column_letter(col_num)
                row_data.append(f"Col {col_num}({col_letter}): {str(val)[:60]}")
                header_data[col_num] = val
        
        if row_data:
            for item in row_data:
                print(f"  {item}")
    
    # Maintenant, analyzer les colonnes DATA (row 11 onwards)
    print(f"\nStructure détectée pour {sheet_name}:")
    print("-" * 150)
    
    # Vérifier quels colonnes contiennent des labels de comptes
    row_11_data = []
    for col_num in range(1, 16):
        cell = ws.cell(11, col_num)
        val = cell.value
        if val:
            col_letter = get_column_letter(col_num)
            row_11_data.append(f"Col {col_num}({col_letter}): {str(val)[:40]}")
    
    if row_11_data:
        print("\nPremière ligne de données (Row 11):")
        for item in row_11_data:
            print(f"  {item}")

print("\n" + "=" * 150)
print("ANALYSE - CE QUE JE REMARQUE")
print("=" * 150)

print("""
Patterns à détecter:

1. COLONNES 'HEADER ONLY' (contiennent des titres, pas de données):
   - Habituellement colonnes 1-3 (pour les libellés de comptes)
   - Ces colonnes ne doivent PAS être remplies

2. COLONNES 'DATA' (contiennent les libellés qu'on cherche):
   - Col 4, 5, 6, 7, 8, 9, 10 probablement
   - Ce sont celles-ci qu'on doit analyser pour mapping

3. COLONNES 'VIDES' (merged ou structurales):
   - À skiper complètement

4. PATTERNS COMMUNS:
   - Libellés "BRUT" = soldes (ouverture/clôture)
   - Libellés "MOUVEMENT", "ACQUI", "CESSION", etc. = détails mouvements
   - Libellés "DEPRECIATION", "AMORT" = dépréciations

5. LOGIQUE COMPTABLE:
   - Vérifie si: Col 10 = Col 4 + (Col 5-9)
   - Si OUI → Colonnes intermédiaires sont DETAILS (besoin détail source)
   - Si NON → Colonnes peut-être indépendantes
""")

wb.close()

print("\n" + "=" * 150)
print("MAINTENANT, Lire tous les libellés spécifiquement")
print("=" * 150)

# Relancer pour affichage complet
wb = openpyxl.load_workbook(dsf_template, data_only=False)

for sheet_name in ["NOTE 3A"]:
    if sheet_name not in wb.sheetnames:
        continue
    
    ws = wb[sheet_name]
    
    print(f"\nDETAIL COMPLET - {sheet_name}:")
    print("-" * 150)
    
    # Afficher rows 8-10 avec numéros de colonne
    print("\nHIERARCHIE DE TITRES:")
    for row_num in [8, 9, 10]:
        print(f"\n  Row {row_num}:")
        for col_num in range(1, 16):
            cell = ws.cell(row_num, col_num)
            val = cell.value
            col_letter = get_column_letter(col_num)
            # Afficher même si vide, pour voir la structure
            val_str = str(val)[:70] if val else "[VIDE]"
            print(f"    {col_letter:3s} (Col {col_num:2d}): {val_str}")

wb.close()
