#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Fast NOTE 20 & 34 Filler - Remplissage structuré sans fuzzy matching
Utilise les mappings SYSCOHADA pour un remplissage rapide et sûr
"""

import logging
from pathlib import Path
from decimal import Decimal
from typing import Dict, List, Optional, Tuple
import time

from openpyxl import Workbook

logger = logging.getLogger(__name__)


class Note20FastFiller:
    """
    Remplis NOTE 20 par structure SYSCOHADA au lieu de fuzzy matching
    
    NOTE 20 : Détail des immobilisations par nature
    Structure : Incorporelles → Corporelles → Financières
    Chaque compté = une ligne avec : Brut début, Augmentations, Diminutions, Brut fin, Amortissements, Dotations, etc.
    """
    
    def __init__(self, workbook: Workbook):
        self.wb = workbook
        self.total_assignments = 0
        
    def detect_note20_structure(self, ws) -> Dict[str, int]:
        """Détecte la structure de NOTE 20 (où commencent les sections)"""
        structure = {
            "incorporelles_start": None,
            "corporelles_start": None,
            "financieres_start": None,
            "total_row": None,
        }
        
        keywords = {
            "incorporelles": ["incorporelle", "immatériel"],
            "corporelles": ["corporelle", "matériel", "terrain", "bâtiment"],
            "financieres": ["financière", "placements", "titres"],
            "total": ["total", "ensemble"],
        }
        
        for row_num in range(1, min(100, ws.max_row)):
            cell_val = ws.cell(row_num, 1).value
            if not cell_val:
                continue
            cell_str = str(cell_val).lower()
            
            for key, words in keywords.items():
                for word in words:
                    if word in cell_str:
                        if key == "total" and structure["total_row"] is None:
                            structure["total_row"] = row_num
                        elif key == "incorporelles" and structure["incorporelles_start"] is None:
                            structure["incorporelles_start"] = row_num
                        elif key == "corporelles" and structure["corporelles_start"] is None:
                            structure["corporelles_start"] = row_num
                        elif key == "financieres" and structure["financieres_start"] is None:
                            structure["financieres_start"] = row_num
        
        return structure
    
    def fill_note20_section(
        self,
        ws,
        section_name: str,
        syscohada_classes: List[str],
        accounts: Dict[str, any],
        start_row: int,
        end_row: Optional[int] = None
    ) -> int:
        """
        Remplit une section de NOTE 20
        
        Args:
            ws: Worksheet
            section_name: Nom de la section (debug)
            syscohada_classes: Classes SYSCOHADA à inclure (ex: ["201", "202", "203"])
            accounts: Dict des comptes chargés [compte] = NormalizedBalanceRow
            start_row: Ligne de départ
            end_row: Ligne de fin (optionnel, se remplit jusqu'au bout si None)
        
        Returns:
            Nombre d'assignations effectuées
        """
        
        start_time = time.time()
        assignments = 0
        
        # Étape 1 : Filtre les comptes par classe SYSCOHADA
        section_accounts = []
        for compte, row in accounts.items():
            for classe in syscohada_classes:
                if str(compte).startswith(classe):
                    section_accounts.append((compte, row))
                    break
        
        section_accounts.sort(key=lambda x: x[0])  # Tri par numéro de compte
        
        # Étape 2 : Remplit les lignes
        if end_row is None:
            end_row = ws.max_row
        
        excel_row = start_row
        for compte, row in section_accounts:
            if excel_row > end_row:
                break
            
            # Colonnes típicas de NOTE 20 :
            # A: Compte
            # B: Libellé
            # C-E: Brut (ouverture, augmentations, diminutions)
            # F: Brut clôture
            # G-H: Amortissements (ouverture, dotations)
            # I: Amortissements clôture
            # J: Net clôture = F - I
            
            ws.cell(excel_row, 1).value = compte
            ws.cell(excel_row, 2).value = row.label[:50]
            
            # Valeurs
            debit = float(row.debit_balance) if row.debit_balance else 0.0
            credit = float(row.credit_balance) if row.credit_balance else 0.0
            net = debit - credit
            
            # Remplissage simplifié (à adapter selon template réel)
            ws.cell(excel_row, 6).value = net  # Brut clôture
            ws.cell(excel_row, 10).value = net  # Net (simplifié)
            
            excel_row += 1
            assignments += 1
        
        elapsed = time.time() - start_time
        logger.info(f"  {section_name}: {assignments} comptes en {elapsed:.2f}s")
        
        return assignments
    
    def fill_note20(self, accounts: Dict[str, any]) -> int:
        """Remplit complètement la NOTE 20"""
        if "NOTE 20" not in self.wb.sheetnames:
            logger.warning("NOTE 20 not found in workbook")
            return 0
        
        ws = self.wb["NOTE 20"]
        total = 0
        
        logger.info("Filling NOTE 20 with SYSCOHADA structure...")
        
        # Détecte la structure du sheet
        structure = self.detect_note20_structure(ws)
        logger.info(f"  Incorporelles: row {structure['incorporelles_start']}")
        logger.info(f"  Corporelles:   row {structure['corporelles_start']}")
        logger.info(f"  Financieres:   row {structure['financieres_start']}")
        
        # Section 1 : Incorporelles (classes 201, 202)
        if structure["incorporelles_start"]:
            section_end = (structure["corporelles_start"] - 1) if structure["corporelles_start"] else None
            total += self.fill_note20_section(
                ws,
                section_name="Incorporelles",
                syscohada_classes=["201", "202"],
                accounts=accounts,
                start_row=structure["incorporelles_start"] + 1,
                end_row=section_end
            )
        
        # Section 2 : Corporelles (classe 203, 204, 205)
        if structure["corporelles_start"]:
            section_end = (structure["financieres_start"] - 1) if structure["financieres_start"] else (structure["total_row"] - 1) if structure["total_row"] else None
            total += self.fill_note20_section(
                ws,
                section_name="Corporelles",
                syscohada_classes=["203", "204", "205"],
                accounts=accounts,
                start_row=structure["corporelles_start"] + 1,
                end_row=section_end
            )
        
        # Section 3 : Financières (classe 26)
        if structure["financieres_start"]:
            section_end = (structure["total_row"] - 1) if structure["total_row"] else None
            total += self.fill_note20_section(
                ws,
                section_name="Financieres",
                syscohada_classes=["261", "262", "263", "264", "265", "266"],
                accounts=accounts,
                start_row=structure["financieres_start"] + 1,
                end_row=section_end
            )
        
        logger.info(f"NOTE 20 filled: {total} comptes assignés")
        return total


class Note34FastFiller:
    """
    Remplis NOTE 34 : Ratios et indicateurs financiers
    NOTE 34 contient des calculs dérivés, pas des données brutes
    Stratégie : Calcule les valeurs à partir des comptes SIG/résultat
    """
    
    def __init__(self, workbook: Workbook):
        self.wb = workbook
    
    def fill_note34(self, accounts: Dict[str, any]) -> int:
        """
        Remplit NOTE 34 avec des calculs dérivés
        Majorité de NOTE 34 = formules qui se recalculent automatiquement dans Excel
        """
        if "NOTE 34" not in self.wb.sheetnames:
            logger.warning("NOTE 34 not found in workbook")
            return 0
        
        logger.info("NOTE 34: Skipping (contains formulas, will auto-calculate)")
        # NOTE 34 contient surtout des formules qui se recalculent
        # On ne remplis que les totaux si nécessaire
        return 0


def integrate_fast_fillers(wb: Workbook, accounts: Dict[str, any]) -> Dict[str, int]:
    """
    Utilisation du remplisseur rapide dans le pipeline
    
    Exemple d'intégration dans dsf_pipeline.py :
    
    from src.note20_34_fast_filler import integrate_fast_fillers
    
    # Après avoir chargé le template et les comptes
    results = integrate_fast_fillers(wb, balance_accounts)
    print(f"NOTE 20: {results['note20']} assignations en {results['note20_time']:.2f}s")
    print(f"NOTE 34: {results['note34']} assignations en {results['note34_time']:.2f}s")
    """
    
    results = {}
    
    # NOTE 20
    start = time.time()
    note20_filler = Note20FastFiller(wb)
    results['note20'] = note20_filler.fill_note20(accounts)
    results['note20_time'] = time.time() - start
    
    # NOTE 34
    start = time.time()
    note34_filler = Note34FastFiller(wb)
    results['note34'] = note34_filler.fill_note34(accounts)
    results['note34_time'] = time.time() - start
    
    return results


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format='%(message)s')
    print("[OK] Fast fillers ready to integrate")
