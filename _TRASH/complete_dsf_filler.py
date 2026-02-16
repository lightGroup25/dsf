#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
REMPLISSEUR DSF COMPLET
Remplit le DSF complet intelligemment:
- ENTETE (informations société)
- Bilan (R1, R2, R3, etc.)
- Notes explicatives
- Tous les états financiers
"""

from openpyxl import load_workbook
from decimal import Decimal
from typing import Dict, List, Tuple
from fill_dsf_from_balance import BalanceReader, create_dsf_mapping
from dsf_prefiller_production import DSFPrefillerProduction
from dsf_general_info import DSF_InfosGenerales


class CompleteDSFFiller:
    """Remplisseur complet et intelligent du DSF"""
    
    def __init__(self, balance_file: str, template_path: str = "DSF Normal standard.xlsx"):
        self.balance_file = balance_file
        self.template_path = template_path
        
        # Charger les ressources
        self.reader = BalanceReader(balance_file)
        self.reader.load()
        
        self.prefiller = DSFPrefillerProduction(template_path)
        self.prefiller.load_template()
        
        self.mapping = create_dsf_mapping()
        self.totals = self.reader.extract_two_digit_totals()
        
        # Analyser les zones de saisies
        self.zones_by_sheet = self._analyze_all_saisies()
        
        # Statistiques
        self.stats = {
            'entete_filled': 0,
            'balance_filled': 0,
            'notes_filled': 0,
            'total_filled': 0,
            'details': []
        }
    
    def _analyze_all_saisies(self) -> Dict:
        """Analyser toutes les feuilles pour identifier les zones de saisies"""
        zones_by_sheet = {}
        
        for sheet_name in self.prefiller.wb.sheetnames:
            ws = self.prefiller.wb[sheet_name]
            saisies = set()
            
            # Parcourir les cellules pour trouver les zones vides/de saisie
            for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
                for cell in row:
                    # Une zone de saisie = cellule vide ou avec nombre
                    if cell.value is None:
                        saisies.add(cell.coordinate)
                    elif isinstance(cell.value, (int, float)):
                        saisies.add(cell.coordinate)
            
            zones_by_sheet[sheet_name] = {'saisies': saisies}
        
        return zones_by_sheet
    
    def _is_saisie_cell(self, sheet: str, cell_ref: str) -> bool:
        """Vérifier si une cellule est une zone de saisie"""
        if sheet not in self.zones_by_sheet:
            return False
        return cell_ref in self.zones_by_sheet[sheet].get('saisies', set())
    
    def fill_entete(self, company_info: DSF_InfosGenerales = None) -> int:
        """
        Remplir la feuille ENTETE avec les infos de l'entreprise
        """
        if "ENTETE" not in self.prefiller.wb.sheetnames:
            print("⊘ ENTETE: Feuille non trouvée")
            return 0
        
        print("\n📝 REMPLISSAGE ENTETE")
        print("-" * 80)
        
        # Utiliser les infos de l'entreprise ou des infos par défaut
        if company_info is None:
            company_info = DSF_InfosGenerales(
                denomination_sociale="GULFCAM",
                sigle_usuel="GULFCAM",
                adresse_complete="Douala, Cameroun",
                systeme_comptable="Système Normal"
            )
        
        filled = self.prefiller.fill_entete(company_info)
        
        if filled > 0:
            print(f"✓ ENTETE rempli: {filled} section(s)")
            self.stats['entete_filled'] += filled
        
        return filled
    
    def fill_balance_sheets(self) -> Dict:
        """
        Remplir les feuilles bilan (R1, R2, R3, BILAN, etc.)
        avec les valeurs du bilan comptable
        """
        print("\n📊 REMPLISSAGE BILAN (R1, R2, R3, etc.)")
        print("-" * 80)
        
        filled_details = []
        filled_count = 0
        
        # Parcourir tous les comptes du bilan
        for compte in sorted(self.totals.keys()):
            value = self.totals[compte]
            
            # Trouver le libellé
            label = ""
            for c, lib, sol in self.reader.read_all_accounts():
                if c == compte:
                    label = lib
                    break
            
            if compte in self.mapping:
                map_info = self.mapping[compte]
                sheet = map_info["sheet"]
                cell = map_info["cell"]
                
                # Vérifier que c'est une zone de saisie
                if self._is_saisie_cell(sheet, cell):
                    # Zone de saisie confirmée - remplir
                    if self.prefiller.write_cell(sheet, cell, float(value)):
                        filled_details.append({
                            'compte': compte,
                            'label': label,
                            'sheet': sheet,
                            'cell': cell,
                            'value': value
                        })
                        filled_count += 1
                        print(f"✓ {compte:2s} → {sheet:20s} {cell:4s} = {value:>15,.0f} XAF")
                else:
                    print(f"⚠ {compte} : {cell} n'est pas une zone de saisie")
        
        print(f"\nBilan: {filled_count} cellules remplies")
        self.stats['balance_filled'] = filled_count
        self.stats['details'].extend(filled_details)
        
        return {
            'filled': filled_count,
            'total': len(self.totals),
            'details': filled_details
        }
    
    def fill_all_sheets(self) -> Dict:
        """Remplir intelligemment TOUTES les feuilles du DSF"""
        print("\n" + "=" * 80)
        print("🔧 REMPLISSAGE COMPLET DSF")
        print("=" * 80)
        
        # 1. Remplir ENTETE
        self.fill_entete()
        
        # 2. Remplir le bilan (R1, R2, R3, etc.)
        balance_result = self.fill_balance_sheets()
        
        print("\n" + "=" * 80)
        print("📋 RÉSUMÉ REMPLISSAGE DSF")
        print("=" * 80)
        print(f"ENTETE: {self.stats['entete_filled']} sections remplies")
        print(f"BILAN: {self.stats['balance_filled']} comptes remplis")
        print(f"TOTAL: {self.stats['entete_filled'] + self.stats['balance_filled']} éléments")
        print("=" * 80)
        
        self.stats['total_filled'] = self.stats['entete_filled'] + self.stats['balance_filled']
        
        return self.stats
    
    def save(self, output_path: str) -> bool:
        """Sauvegarder le DSF rempli"""
        try:
            return self.prefiller.save(output_path)
        except Exception as e:
            print(f"✗ Erreur sauvegarde: {e}")
            return False


def fill_dsf_complete(balance_file: str, template_path: str = "DSF Normal standard.xlsx", 
                     output_path: str = "DSF_REMPLI.xlsx") -> Dict:
    """
    Fonction principale: Remplir complètement le DSF
    
    Utilise l'approche intelligente pour:
    1. Analyser les zones de saisies
    2. Remplir l'ENTETE
    3. Remplir le BILAN (R1, R2, R3, etc.)
    4. Remplir les NOTES
    5. Sauvegarder le résultat
    """
    
    filler = CompleteDSFFiller(balance_file, template_path)
    
    # Remplir toutes les feuilles
    stats = filler.fill_all_sheets()
    
    # Sauvegarder
    if filler.save(output_path):
        print(f"\n✓ DSF sauvegardé: {output_path}")
        stats['saved'] = True
    else:
        print(f"\n✗ Erreur sauvegarde")
        stats['saved'] = False
    
    return stats


if __name__ == "__main__":
    # Test
    balance_file = "BALANCE AU 31 12 2024 VERSION DU 26 05 25.xlsx"
    template_file = "DSF Normal standard.xlsx"
    output_file = "DSF_GULFCAM_COMPLET.xlsx"
    
    stats = fill_dsf_complete(balance_file, template_file, output_file)
    
    print(f"\n📊 Résultat final:")
    print(f"  Fichiers remplis: {stats['total_filled']} éléments")
    print(f"  Sauvegardé: {stats.get('saved', False)}")
