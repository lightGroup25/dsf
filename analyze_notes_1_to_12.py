#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Analyse des Notes 1-12 du DSF 2023 pour vérifier leur structure et contenu"""

from openpyxl import load_workbook
from pathlib import Path
import sys

dsf_2023 = Path(r"C:\Users\Emmaneul Ambadiang\Desktop\dsf\input\DSF GULFCAM 2023 V3.xlsx")

if not dsf_2023.exists():
    print(f"[NO] Fichier non trouve: {dsf_2023}")
    sys.exit(1)

try:
    wb = load_workbook(dsf_2023, data_only=True)
    
    print("=" * 100)
    print("ANALYSE DES NOTES 1-12 DU DSF 2023")
    print("=" * 100)
    
    # Verifier les notes 1 a 12
    note_names = [
        "Note 1", "NOTE 2", "NOTE 3A", "NOTE 3B", "NOTE 3C", "NOTE 4",
        "NOTE 5", "NOTE 6", "NOTE 7", "NOTE 8", "NOTE 9", "NOTE 10", 
        "NOTE 11", "NOTE 12"
    ]
    
    print("\n### PRESENCE DES NOTES ANNEXES ###")
    for note_name in note_names:
        # Essayer variantes (majuscules, espaces, etc.)
        found = False
        actual_name = None
        for sheet in wb.sheetnames:
            if sheet.upper() == note_name.upper():
                found = True
                actual_name = sheet
                break
        
        status = "[OK]" if found else "[MISSING]"
        print(f"  {status} {note_name:15} -> {actual_name if actual_name else 'N/A'}")
    
    # Examiner la structure de NOTE 1
    print("\n### STRUCTURE NOTE 1 (si presente) ###")
    if "Note 1" in wb.sheetnames:
        ws = wb["Note 1"]
        print(f"  Max rows: {ws.max_row}, Max cols: {ws.max_column}")
        
        print("\n  Premieres lignes (A:E):")
        for row in range(1, min(8, ws.max_row + 1)):
            line_parts = []
            for col in range(1, 6):
                cell = ws.cell(row, col)
                val = str(cell.value)[:25] if cell.value else ""
                line_parts.append(val)
            
            line_str = " | ".join(f"{v:23}" for v in line_parts)
            print(f"    Row {row:2}: {line_str}")
    
    # Examiner NOTE 3C
    print("\n### STRUCTURE NOTE 3C ###")
    if "NOTE  3C" in wb.sheetnames:
        ws = wb["NOTE  3C"]
        print(f"  Max rows: {ws.max_row}, Max cols: {ws.max_column}")
        
        print("\n  Premieres lignes (A:E):")
        for row in range(1, min(8, ws.max_row + 1)):
            line_parts = []
            for col in range(1, 6):
                cell = ws.cell(row, col)
                val = str(cell.value)[:25] if cell.value else ""
                line_parts.append(val)
            
            line_str = " | ".join(f"{v:23}" for v in line_parts)
            print(f"    Row {row:2}: {line_str}")
    
    # Verifier le mapping
    print("\n### MAPPING DANS dsf_notes_filler.py ###")
    sys.path.insert(0, r"C:\Users\Emmaneul Ambadiang\Desktop\DSF\src")
    
    try:
        from dsf_notes_filler import NOTE_ACCOUNT_MAPPING
        
        for note_key in ["NOTE 1", "NOTE 2", "NOTE 3A", "NOTE 3B", "NOTE 3C", "NOTE 4", "NOTE 5", "NOTE 6"]:
            if note_key in NOTE_ACCOUNT_MAPPING:
                count = len(NOTE_ACCOUNT_MAPPING[note_key])
                print(f"  [OK] {note_key:15} - {count} comptes mapping")
            else:
                print(f"  [MISSING] {note_key:15} - PAS DANS LE MAPPING")
    except Exception as e:
        print(f"  [ERROR] Impossible de charger le mapping: {e}")
    
    print("\nFin!")
    wb.close()
    
except Exception as e:
    print(f"ERROR: {e}")
    import traceback
    traceback.print_exc()

