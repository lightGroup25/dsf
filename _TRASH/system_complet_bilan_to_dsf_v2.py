#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SYSTÈME COMPLET: Bilan GULFCAM → DSF
Version 2.0 - Avec detection de structure correcte
"""

import openpyxl
from openpyxl.utils import get_column_letter
from typing import Dict, Tuple, Generator
from decimal import Decimal

class BalanceReader:
    """Lecteur du fichier Bilan GULFCAM avec structure SAGE"""
    
    def __init__(self, balance_file: str):
        self.balance_file = balance_file
        self.wb = None
        self.ws = None
        self.compte_col = None
        self.libelle_col = None
        self.solde_col = None
        self.data_start_row = None
        
    def load(self):
        """Charger et analyser le fichier bilan"""
        print(f"\n📖 CHARGEMENT BILAN: {self.balance_file}")
        self.wb = openpyxl.load_workbook(self.balance_file, data_only=True)
        self.ws = self.wb.active
        
        # Structure SAGE: 
        # - Col A: Numéro de compte
        # - Col D: Intitulé des comptes
        # - Data starts row 13
        # - Final balance in Col AA (Soldes cumulés)
        
        self.compte_col = 1      # Column A
        self.libelle_col = 4     # Column D
        self.solde_col = 27      # Column AA (last column with balances)
        self.data_start_row = 13
        
        print(f"✓ Fichier chargé:")
        print(f"  - Feuille: {self.ws.title}")
        print(f"  - Colonnes: Compte(A), Libellé(D), Solde(AA)")
        print(f"  - Données à partir de ligne: {self.data_start_row}")
        
        return self
    
    def read_all_accounts(self) -> Generator[Tuple[str, str, Decimal], None, None]:
        """Générateur: lire tous les comptes avec leurs soldes
        
        Yields:
            (compte: str, libelle: str, solde: Decimal)
        """
        max_row = self.ws.max_row
        
        for row_idx in range(self.data_start_row, max_row + 1):
            compte = self.ws.cell(row=row_idx, column=self.compte_col).value
            libelle = self.ws.cell(row=row_idx, column=self.libelle_col).value
            solde = self.ws.cell(row=row_idx, column=self.solde_col).value
            
            # Filtrer les lignes sans compte
            if not compte or compte == '':
                continue
            
            # Nettoyer le compte (supprimer espaces)
            compte_str = str(compte).strip()
            
            # Convertir solde en Decimal
            if solde is None:
                solde = Decimal('0')
            else:
                try:
                    solde = Decimal(str(solde))
                except:
                    solde = Decimal('0')
            
            yield (compte_str, libelle or '', solde)
    
    def extract_two_digit_totals(self) -> Dict[str, Decimal]:
        """Extraire les comptes à 2 chiffres (totaux de classe)
        
        Returns:
            Dict: {"10": 50000000, "20": 12000000, ...}
        """
        totals = {}
        
        for compte, libelle, solde in self.read_all_accounts():
            # Garder que les comptes 2 chiffres
            if len(compte) == 2 and compte.isdigit():
                totals[compte] = solde
                print(f"  ✓ Compte {compte}: {solde:>15,.0f} - {libelle}")
        
        return totals


def create_totals_mapping() -> Dict[str, Dict]:
    """Créer la correspondance comptes 2 chiffres → cellules DSF
    
    Retourne:
        Dict: {"10": {"sheet": "BILAN PAYSAGE", "cell": "D5", ...}, ...}
    """
    # Mapping complet basé sur SYSCOHADA → DSF Normal
    mapping = {
        # AKTIVA (Actif)
        "10": {
            "sheet": "BILAN PAYSAGE", 
            "cell": "B5",
            "type": "asset",
            "label": "CAPITAL SOCIAL"
        },
        "13": {
            "sheet": "BILAN PAYSAGE",
            "cell": "B6", 
            "type": "asset",
            "label": "Primes d'émission"
        },
        "14": {
            "sheet": "BILAN PAYSAGE",
            "cell": "B7",
            "type": "asset", 
            "label": "Écarts de réévaluation"
        },
        "15": {
            "sheet": "BILAN PAYSAGE",
            "cell": "B8",
            "type": "asset",
            "label": "Écarts d'acquisition"
        },
        "16": {
            "sheet": "BILAN PAYSAGE",
            "cell": "B9",
            "type": "asset",
            "label": "Réserves"
        },
        "17": {
            "sheet": "BILAN PAYSAGE",
            "cell": "B10",
            "type": "asset",
            "label": "Report à nouveau"
        },
        "18": {
            "sheet": "BILAN PAYSAGE",
            "cell": "B11",
            "type": "asset",
            "label": "Résultat de l'exercice"
        },
        "19": {
            "sheet": "BILAN PAYSAGE",
            "cell": "B12",
            "type": "asset",
            "label": "Compte de liaison"
        },
        
        # PASSIVA (Passif à long terme)
        "20": {
            "sheet": "BILAN PAYSAGE",
            "cell": "B16",
            "type": "liability",
            "label": "Emprunts obligataires"
        },
        "21": {
            "sheet": "BILAN PAYSAGE",
            "cell": "B17",
            "type": "liability",
            "label": "Dettes financières à LT"
        },
        "23": {
            "sheet": "BILAN PAYSAGE",
            "cell": "B18",
            "type": "liability",
            "label": "Dettes de location-financement"
        },
        "24": {
            "sheet": "BILAN PAYSAGE",
            "cell": "B19",
            "type": "liability",
            "label": "Dettes liées à des participations"
        },
        "25": {
            "sheet": "BILAN PAYSAGE",
            "cell": "B20",
            "type": "liability",
            "label": "Autres dettes à LT"
        },
        "26": {
            "sheet": "BILAN PAYSAGE",
            "cell": "B21",
            "type": "liability",
            "label": "Provisions pour risques"
        },
        "27": {
            "sheet": "BILAN PAYSAGE",
            "cell": "B22",
            "type": "liability",
            "label": "Provisions pour charges"
        },
        
        # AKTIVA (suite - Actif immobilisé)
        "30": {
            "sheet": "BILAN PAYSAGE",
            "cell": "B25",
            "type": "asset",
            "label": "Immobilisations incorporelles"
        },
        "31": {
            "sheet": "BILAN PAYSAGE",
            "cell": "B26",
            "type": "asset",
            "label": "Terrains"
        },
        "32": {
            "sheet": "BILAN PAYSAGE",
            "cell": "B27",
            "type": "asset",
            "label": "Constructions"
        },
        "33": {
            "sheet": "BILAN PAYSAGE",
            "cell": "B28",
            "type": "asset",
            "label": "Inst./Mach./Outillage"
        },
        "34": {
            "sheet": "BILAN PAYSAGE",
            "cell": "B29",
            "type": "asset",
            "label": "Mobilier/Matériel roulant"
        },
        "35": {
            "sheet": "BILAN PAYSAGE",
            "cell": "B30",
            "type": "asset",
            "label": "Avances et acomptes"
        },
        "36": {
            "sheet": "BILAN PAYSAGE",
            "cell": "B31",
            "type": "asset",
            "label": "Immobilisations financières"
        },
        
        # PASSIVA (suite - Passif courant)
        "40": {
            "sheet": "BILAN PAYSAGE",
            "cell": "D15",
            "type": "liability",
            "label": "Créances commerciales"
        },
        "41": {
            "sheet": "BILAN PAYSAGE",
            "cell": "B35",
            "type": "asset",
            "label": "Créances clients"
        },
        "42": {
            "sheet": "BILAN PAYSAGE",
            "cell": "D16",
            "type": "liability",
            "label": "Dettes commerciales"
        },
        "44": {
            "sheet": "BILAN PAYSAGE",
            "cell": "D17",
            "type": "liability",
            "label": "Dettes fiscales"
        },
        "45": {
            "sheet": "BILAN PAYSAGE",
            "cell": "D18",
            "type": "liability",
            "label": "Dettes sociales"
        },
        "47": {
            "sheet": "BILAN PAYSAGE",
            "cell": "D19",
            "type": "liability",
            "label": "Autres dettes"
        },
        
        # AKTIVA (suite - Actif courant)
        "50": {
            "sheet": "BILAN PAYSAGE",
            "cell": "B34",
            "type": "asset",
            "label": "Stocks matières premières"
        },
        "51": {
            "sheet": "BILAN PAYSAGE",
            "cell": "B35",
            "type": "asset",
            "label": "Stocks en-cours"
        },
        "52": {
            "sheet": "BILAN PAYSAGE",
            "cell": "B36",
            "type": "asset",
            "label": "Stocks produits finis"
        },
        "55": {
            "sheet": "BILAN PAYSAGE",
            "cell": "B37",
            "type": "asset",
            "label": "Stocks marchandises"
        },
        
        # PASSIVA (suite - Dettes financières CT)
        "50": {
            "sheet": "BILAN PAYSAGE",
            "cell": "D13",
            "type": "liability",
            "label": "Emprunts bancaires CT"
        },
        
        # AKTIVA (Trésorerie-actif)
        "57": {
            "sheet": "BILAN PAYSAGE",
            "cell": "B40",
            "type": "asset",
            "label": "Titres de placement"
        },
        "58": {
            "sheet": "BILAN PAYSAGE",
            "cell": "B41",
            "type": "asset",
            "label": "Écarts de conversion-actif"
        },
        "59": {
            "sheet": "BILAN PAYSAGE",
            "cell": "B42",
            "type": "asset",
            "label": "Banques/Chèques/Caisse"
        },
        
        # COMPTE RESULTAT - CHARGES (60, 61, 62...)
        "60": {
            "sheet": "COMPTE RESULTAT",
            "cell": "D8",
            "type": "expense",
            "label": "Achats"
        },
        "61": {
            "sheet": "COMPTE RESULTAT",
            "cell": "D9",
            "type": "expense",
            "label": "Variation de stocks"
        },
        "62": {
            "sheet": "COMPTE RESULTAT",
            "cell": "D10",
            "type": "expense",
            "label": "Transports"
        },
        "63": {
            "sheet": "COMPTE RESULTAT",
            "cell": "D11",
            "type": "expense",
            "label": "Services extérieurs"
        },
        "64": {
            "sheet": "COMPTE RESULTAT",
            "cell": "D12",
            "type": "expense",
            "label": "Autres services extérieurs"
        },
        "65": {
            "sheet": "COMPTE RESULTAT",
            "cell": "D13",
            "type": "expense",
            "label": "Charges de personnel"
        },
        "66": {
            "sheet": "COMPTE RESULTAT",
            "cell": "D14",
            "type": "expense",
            "label": "Frais de location et charges"
        },
        "67": {
            "sheet": "COMPTE RESULTAT",
            "cell": "D15",
            "type": "expense",
            "label": "Droits et taxes"
        },
        "68": {
            "sheet": "COMPTE RESULTAT",
            "cell": "D16",
            "type": "expense",
            "label": "Autres charges"
        },
        
        # COMPTE RESULTAT - PRODUITS (70, 71...)
        "70": {
            "sheet": "COMPTE RESULTAT",
            "cell": "C8",
            "type": "income",
            "label": "Ventes de biens et services"
        },
        "71": {
            "sheet": "COMPTE RESULTAT",
            "cell": "C9",
            "type": "income",
            "label": "Variation de stocks produits"
        },
        "72": {
            "sheet": "COMPTE RESULTAT",
            "cell": "C10",
            "type": "income",
            "label": "Produits accessoires"
        },
    }
    
    return mapping


def fill_dsf_from_balance(balance_totals: Dict[str, Decimal], mapping: Dict[str, Dict]):
    """Remplir le DSF à partir des totaux de bilan
    
    Args:
        balance_totals: Dict des comptes 2-chiffres et leurs soldes
        mapping: Correspondance comptes → cellules DSF
    """
    import os
    
    template_file = "DSF Normal standard.xlsx"
    output_file = "DSF_GULFCAM_FILLED.xlsx"
    
    if not os.path.exists(template_file):
        print(f"❌ Template non trouvé: {template_file}")
        return
    
    print(f"\n📝 REMPLISSAGE DSF:")
    print(f"  Template: {template_file}")
    print(f"  Sortie: {output_file}")
    
    # Importer le prefiller
    from dsf_prefiller_production import DSFPrefillerProduction
    
    prefiller = DSFPrefillerProduction(template_file)
    prefiller.load_template()
    
    # Remplir les cellules basées sur les comptes 2-chiffres
    filled_count = 0
    for compte, solde in balance_totals.items():
        if compte in mapping:
            cell_info = mapping[compte]
            sheet = cell_info.get("sheet")
            cell = cell_info.get("cell")
            label = cell_info.get("label", "")
            
            try:
                prefiller.write_cell(sheet, cell, float(solde))
                print(f"  ✓ {compte} → {sheet}!{cell}: {solde:>15,.0f} ({label})")
                filled_count += 1
            except Exception as e:
                print(f"  ⚠ {compte} → {sheet}!{cell}: Erreur - {e}")
        else:
            print(f"  - {compte}: Pas de mapping DSF")
    
    # Sauvegarder
    prefiller.save(output_file)
    print(f"\n✅ {filled_count} cellules remplies")
    print(f"📄 DSF sauvegardé: {output_file}")
    
    return output_file


def main():
    """Pipeline complète: Bilan → DSF"""
    
    print("=" * 80)
    print("SYSTÈME COMPLET: Bilan GULFCAM → DSF (v2.0)")
    print("=" * 80)
    
    balance_file = "BALANCE AU 31 12 2024 VERSION DU 26 05 25.xlsx"
    
    try:
        # 1. Lire le bilan
        reader = BalanceReader(balance_file)
        reader.load()
        
        # 2. Extraire les totaux à 2 chiffres
        print(f"\n📊 TOTAUX DE BILAN (comptes 2-chiffres):")
        totals = reader.extract_two_digit_totals()
        print(f"Total: {len(totals)} comptes extraits")
        
        if not totals:
            print("⚠️ Aucun compte 2-chiffres trouvé!")
            return
        
        # 3. Créer le mapping
        print(f"\n🔗 CORRESPONDANCE COMPTES → DSF:")
        mapping = create_totals_mapping()
        print(f"Total: {len(mapping)} correspondances disponibles")
        
        # 4. Remplir le DSF
        fill_dsf_from_balance(totals, mapping)
        
        print("\n✅ PIPELINE COMPLÈTE TERMINÉE")
        
    except Exception as e:
        import traceback
        print(f"\n❌ ERREUR: {e}")
        traceback.print_exc()


if __name__ == "__main__":
    main()
