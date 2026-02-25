"""
Analyze NOTE 3A CELLULE PAR CELLULE pour comprendre la structure exacte
- Lire les libellés des colonnes (pas supposer)
- Lire les libellés des lignes
- Déterminer quelles colonnes doivent être remplies
- Déterminer quelle donnée balance va dans quelle colonne
"""

import openpyxl

dsf_file = r'templates/DSF Normal standard.xlsx'
wb = openpyxl.load_workbook(dsf_file)

# Analyser NOTE 3A ligne par ligne, colonne par colonne
ws = wb['NOTE 3A']

print("="*70)
print("ANALYSE NOTE 3A - STRUCTURE EXACTE")
print("="*70)

# D'abord, afficher les en-têtes (structure des colonnes)
print("\nEN-TÊTES DES COLONNES (Rows 1-10):")
print("-" * 70)

for row_idx in range(1, 12):
    row_content = []
    for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']:
        cell = ws[f'{col}{row_idx}']
        val = cell.value
        if val:
            row_content.append(f"{col}: {str(val)[:30]}")
    
    if row_content:
        print(f"Row {row_idx}: {', '.join(row_content)}")

print("\n" + "="*70)
print("PREMIÈRES LIGNES DE DONNÉES (Rows 12-25):")
print("-" * 70)

for row_idx in range(12, 26):
    col_a = ws[f'A{row_idx}'].value
    col_b = ws[f'B{row_idx}'].value
    col_d = ws[f'D{row_idx}'].value
    col_e = ws[f'E{row_idx}'].value
    col_f = ws[f'F{row_idx}'].value
    col_k = ws[f'K{row_idx}'].value
    
    if col_a or col_b:
        print(f"\nRow {row_idx}:")
        print(f"  A: {col_a}")
        print(f"  B: {col_b}")
        print(f"  D: {col_d}")
        print(f"  E: {col_e}")
        print(f"  F: {col_f}")
        print(f"  K: {col_k}")

print("\n" + "="*70)
print("COLONNES PRÉSENTES AVEC DONNÉES:")
print("-" * 70)

# Scanner toutes les colonnes utilisées
columns_used = set()
for row_idx in range(1, min(100, ws.max_row + 1)):
    for col_letter in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']:
        cell = ws[f'{col_letter}{row_idx}']
        if cell.value:
            columns_used.add(col_letter)

print(f"Colonnes avec contenu: {sorted(columns_used)}")

# Montrer les libellés de chaque colonne utilisée
print("\nLibellés des colonnes utilisées:")
for col in sorted(columns_used):
    # Trouver le libellé de la colonne (généralement row 5-10)
    for row_idx in range(1, 15):
        cell_val = ws[f'{col}{row_idx}'].value
        if cell_val and 'BILAN' not in str(cell_val) and 'BRUT' in str(cell_val).upper():
            print(f"  Col {col}: {cell_val}")
            break
        elif cell_val and any(x in str(cell_val).upper() for x in ['MONTANT', 'VALEUR', 'EXERCICE', 'OUVERTURE', 'CLOTURE', 'NET']):
            print(f"  Col {col}: {cell_val}")
            break
