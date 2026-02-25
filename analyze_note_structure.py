#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Analyseur de structure NOTE 20 et 34 - Extraire le mapping DSF → SYSCOHADA
"""

import sys
from pathlib import Path

sys.path.insert(0, 'src')

from openpyxl import load_workbook


def _safe_text(value: object) -> str:
    text = str(value)
    return text.encode("ascii", "ignore").decode("ascii")

template_path = Path("templates/DSF Normal standard.xlsx")
if not template_path.exists():
    print(f"[!] Template non trouvé: {template_path}")
    print("    Cherchons les fichiers Excel disponibles...")
    for f in Path("input").glob("*.xlsx"):
        print(f"    Trouvé: {f.name}")
    sys.exit(1)

wb = load_workbook(template_path, data_only=False)

target_notes = [
    "NOTE 18",
    "NOTE 19",
    "NOTE 20",
    "NOTE 27A",
    "NOTE 28",
    "C1-NOTE 28",
    "C2-NOTE 28",
    "NOTE 30",
    "NOTE 32",
    "NOTE 15A",
    "NOTE 15B",
    "NOTE 16A",
    "C1-NOTE 17",
    "C1-NOTE 25",
    "C2-NOTE 25",
    "NOTE 34",
]

print("\n" + "="*80)
print("ANALYSE STRUCTURE NOTES CIBLEES")
print("="*80)

sheet_name_map = {" ".join(name.upper().split()): name for name in wb.sheetnames}

for sheet_name in target_notes:
    normalized = " ".join(sheet_name.upper().split())
    actual_name = sheet_name_map.get(normalized)
    if not actual_name:
        print(f"\n[!] {sheet_name} non trouvée")
        continue
    
    ws = wb[actual_name]
    print(f"\n{sheet_name}:")
    print(f"  Dimensions: {ws.dimensions}")
    print(f"  Merged cells: {len(ws.merged_cells.ranges)}")
    
    # Extraire les labels de lignes (colonne A généralement)
    print(f"\n  Premières lignes (col A = labels):")
    row_labels = {}
    for row_num in range(1, min(50, ws.max_row + 1)):
        cell_val = ws.cell(row_num, 1).value
        if cell_val:
            cell_val_str = str(cell_val).strip()
            if len(cell_val_str) > 0 and not cell_val_str.startswith(" "):
                row_labels[row_num] = cell_val_str
                if len(row_labels) <= 20:
                    print(f"    Ligne {row_num:3d}: {_safe_text(cell_val_str)[:70]}")
    
    # Extraire les labels de colonnes (lignes 1-5 pour trouver l'entete)
    print(f"\n  En-têtes colonnes (lignes 1-5):")
    for header_row in range(1, 6):
        col_labels = {}
        for col_num in range(1, min(15, ws.max_column + 1)):
            cell_val = ws.cell(header_row, col_num).value
            if cell_val:
                col_labels[col_num] = _safe_text(cell_val).strip()[:40]
        if col_labels:
            print(f"    Ligne {header_row}: {col_labels}")
    
    # Identifier la ligne d'en-tetes autour de "Libelle"
    header_row = None
    for row_num, label in sorted(row_labels.items()):
        label_norm = _safe_text(label).lower()
        if "libell" in label_norm:
            header_row = row_num
            break
    if header_row:
        print(f"\n  En-tetes detectes a la ligne {header_row}:")
        header_cells = {}
        for col_num in range(1, min(15, ws.max_column + 1)):
            cell_val = ws.cell(header_row, col_num).value
            if cell_val:
                header_cells[col_num] = _safe_text(cell_val).strip()[:40]
        print(f"    {header_cells}")

    # Identifier les sections principales
    print(f"\n  Sections principales:")
    current_section = None
    for row_num, label in sorted(row_labels.items()):
        if row_num <= 10 or "TOTAL" in label or "total" in label.lower():
            if label != current_section:
                print(f"    [{row_num:3d}] {_safe_text(label)[:60]}")
                current_section = label

print("\n" + "="*80)
print("Feuilles disponibles dans le template:")
print("="*80)
for i, sheet in enumerate(wb.sheetnames, 1):
    ws = wb[sheet]
    print(f"{i:2d}. {sheet:20s} - Dim: {ws.dimensions}")

wb.close()
print("\nAnalyse terminée.")
