#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SYSTÈME COMPLET: GULFCAM Bilan → DSF
Version Finale - Remplissage DSF à partir du bilan de comptes

Ce système:
1. Lit le fichier bilan GULFCAM (Balance des comptes)
2. Extrait les comptes à 2 chiffres (totaux de classe)
3. Remplit le DSF avec ces totaux
4. Sauvegarde le DSF complété

Structure du bilan GULFCAM (fichier SAGE):
- Ligne 1-12: En-têtes informationnels
- Ligne 8, Col A: "Numéro de compte"
- Ligne 9, Col F: "Intitulé des comptes"
- Ligne 13+: Données des comptes
- Col A: Numéro de compte
- Col D: Libellé du compte
- Col AA: Solde cumulé (final balance)
"""

import openpyxl
from typing import Dict, Tuple, Generator
from decimal import Decimal
import os


class BalanceReader:
    """Lecteur du fichier Bilan GULFCAM (structure SAGE)"""
    
    def __init__(self, balance_file: str):
        self.balance_file = balance_file
        self.wb = None
        self.ws = None
        # Structure fixe du fichier SAGE Balance GULFCAM
        self.compte_col = 1      # Column A
        self.libelle_col = 4     # Column D
        self.solde_col = 27      # Column AA
        self.data_start_row = 13
        
    def load(self):
        """Charger le fichier bilan"""
        if not os.path.exists(self.balance_file):
            raise FileNotFoundError(f"Bilan non trouvé: {self.balance_file}")
        
        print(f"📖 CHARGEMENT: {self.balance_file}")
        self.wb = openpyxl.load_workbook(self.balance_file, data_only=True)
        self.ws = self.wb.active
        
        print(f"✓ Feuille active: {self.ws.title}")
        print(f"✓ Structure: Compte(A) | Libellé(D) | Solde(AA)")
        print(f"✓ Données à partir de ligne {self.data_start_row}")
        return self
    
    def read_all_accounts(self) -> Generator[Tuple[str, str, Decimal], None, None]:
        """Générateur des comptes avec leurs soldes
        
        Yields:
            (compte: str, libelle: str, solde: Decimal)
        """
        max_row = self.ws.max_row
        
        for row_idx in range(self.data_start_row, max_row + 1):
            compte = self.ws.cell(row=row_idx, column=self.compte_col).value
            libelle = self.ws.cell(row=row_idx, column=self.libelle_col).value
            solde = self.ws.cell(row=row_idx, column=self.solde_col).value
            
            if not compte:
                continue
            
            compte_str = str(compte).strip()
            
            if solde is None:
                solde = Decimal('0')
            else:
                try:
                    solde = Decimal(str(solde))
                except:
                    solde = Decimal('0')
            
            yield (compte_str, libelle or '', solde)
    
    def extract_two_digit_totals(self) -> Dict[str, Decimal]:
        """Extraire les comptes à 2 chiffres (totaux de classe SYSCOHADA)
        
        Returns:
            Dict: {"10": 12049982015, "70": 18447136737, ...}
        """
        totals = {}
        
        for compte, libelle, solde in self.read_all_accounts():
            if len(compte) == 2 and compte.isdigit():
                totals[compte] = solde
        
        return totals


def create_dsf_mapping() -> Dict[str, Dict]:
    """Créer mapping comptes SYSCOHADA 2-chiffres → cellules DSF
    
    Note: Ce mapping est basé sur la correspondance SYSCOHADA → DSF Normal
    Les numéros de cellules correspondent au fichier "DSF Normal standard.xlsx"
    """
    mapping = {
        # Équité / Passif à LT
        "10": {"sheet": "BILAN PAYSAGE", "cell": "C5", "label": "Capital social"},
        "11": {"sheet": "BILAN PAYSAGE", "cell": "C6", "label": "Réserves"},
        "12": {"sheet": "BILAN PAYSAGE", "cell": "C7", "label": "Report à nouveau"},
        "13": {"sheet": "BILAN PAYSAGE", "cell": "C8", "label": "Résultat net"},
        "16": {"sheet": "BILAN PAYSAGE", "cell": "C11", "label": "Emprunts L.T."},
        "17": {"sheet": "BILAN PAYSAGE", "cell": "C12", "label": "Dettes location-financement"},
        "19": {"sheet": "BILAN PAYSAGE", "cell": "C14", "label": "Provisions"},
        
        # Actif immobilisé
        "21": {"sheet": "BILAN PAYSAGE", "cell": "B9", "label": "Incorporelles"},
        "22": {"sheet": "BILAN PAYSAGE", "cell": "B10", "label": "Terrains"},
        "23": {"sheet": "BILAN PAYSAGE", "cell": "B11", "label": "Constructions"},
        "24": {"sheet": "BILAN PAYSAGE", "cell": "B12", "label": "Matériel"},
        "25": {"sheet": "BILAN PAYSAGE", "cell": "B13", "label": "Avances/acomptes"},
        "26": {"sheet": "BILAN PAYSAGE", "cell": "B14", "label": "Participations"},
        "27": {"sheet": "BILAN PAYSAGE", "cell": "B15", "label": "Autres immo fin."},
        
        # Dépréciations et amortissements
        "28": {"sheet": "BILAN PAYSAGE", "cell": "B18", "label": "Amortissements"},
        "29": {"sheet": "BILAN PAYSAGE", "cell": "B19", "label": "Dépréciations"},
        
        # Stocks
        "31": {"sheet": "BILAN PAYSAGE", "cell": "B21", "label": "Stocks marchandises"},
        "33": {"sheet": "BILAN PAYSAGE", "cell": "B22", "label": "Approvisionnements"},
        
        # Fournisseurs et dettes CT
        "40": {"sheet": "BILAN PAYSAGE", "cell": "D5", "label": "Fournisseurs"},
        "41": {"sheet": "BILAN PAYSAGE", "cell": "B25", "label": "Clients"},
        "42": {"sheet": "BILAN PAYSAGE", "cell": "D6", "label": "Personnel"},
        "43": {"sheet": "BILAN PAYSAGE", "cell": "D7", "label": "Organismes sociaux"},
        "44": {"sheet": "BILAN PAYSAGE", "cell": "D8", "label": "Fiscalité"},
        "45": {"sheet": "BILAN PAYSAGE", "cell": "D9", "label": "Dettes sociales"},
        "47": {"sheet": "BILAN PAYSAGE", "cell": "D11", "label": "Autres dettes"},
        
        # Trésorerie et autres
        "52": {"sheet": "BILAN PAYSAGE", "cell": "B29", "label": "Banques"},
        "55": {"sheet": "BILAN PAYSAGE", "cell": "B30", "label": "Instruments électroniques"},
        "56": {"sheet": "BILAN PAYSAGE", "cell": "B31", "label": "Banques/crédits"},
        "57": {"sheet": "BILAN PAYSAGE", "cell": "B32", "label": "Caisse"},
        
        # Compte résultat - Charges
        "60": {"sheet": "COMPTE RESULTAT", "cell": "D8", "label": "Achats"},
        "61": {"sheet": "COMPTE RESULTAT", "cell": "D9", "label": "Variation stocks"},
        "62": {"sheet": "COMPTE RESULTAT", "cell": "D10", "label": "Transports"},
        "63": {"sheet": "COMPTE RESULTAT", "cell": "D11", "label": "Services ext. 1"},
        "64": {"sheet": "COMPTE RESULTAT", "cell": "D12", "label": "Services ext. 2"},
        "65": {"sheet": "COMPTE RESULTAT", "cell": "D13", "label": "Charges personnel"},
        "66": {"sheet": "COMPTE RESULTAT", "cell": "D14", "label": "Location/charges"},
        "67": {"sheet": "COMPTE RESULTAT", "cell": "D15", "label": "Droits/taxes"},
        "68": {"sheet": "COMPTE RESULTAT", "cell": "D16", "label": "Autres charges"},
        
        # Compte résultat - Produits
        "70": {"sheet": "COMPTE RESULTAT", "cell": "C8", "label": "Ventes"},
        "71": {"sheet": "COMPTE RESULTAT", "cell": "C9", "label": "Var. stocks prod."},
        "72": {"sheet": "COMPTE RESULTAT", "cell": "C10", "label": "Produits access."},
        "73": {"sheet": "COMPTE RESULTAT", "cell": "C11", "label": "Retours/rabais"},
        "75": {"sheet": "COMPTE RESULTAT", "cell": "C13", "label": "Autres produits"},
    }
    
    return mapping


def fill_dsf_template(balance_totals: Dict[str, Decimal], mapping: Dict[str, Dict]):
    """Remplir le template DSF avec les totaux du bilan
    
    Args:
        balance_totals: Dict des comptes → soldes
        mapping: Dict de correspondance comptes → cellules DSF
    """
    from dsf_prefiller_production import DSFPrefillerProduction
    
    template_file = "DSF Normal standard.xlsx"
    output_file = "DSF_GULFCAM_FILLED.xlsx"
    
    if not os.path.exists(template_file):
        print(f"❌ Template non trouvé: {template_file}")
        return
    
    print(f"\n📝 REMPLISSAGE DSF:")
    print(f"   Template: {template_file}")
    print(f"   Sortie: {output_file}")
    
    prefiller = DSFPrefillerProduction(template_file)
    prefiller.load_template()
    
    # Compter les remplissages
    filled = 0
    skipped = 0
    
    for compte, solde in sorted(balance_totals.items()):
        if compte in mapping:
            info = mapping[compte]
            sheet = info["sheet"]
            cell = info["cell"]
            label = info["label"]
            
            try:
                prefiller.write_cell(sheet, cell, float(solde))
                filled += 1
                status = "✓"
            except Exception as e:
                status = "⚠"
                skipped += 1
            
            print(f"   {status} C{compte}: {sheet}!{cell} = {solde:>15,.0f}")
        else:
            print(f"   - C{compte}: Pas de mapping")
            skipped += 1
    
    # Sauvegarder
    prefiller.save(output_file)
    
    print(f"\n✅ Résumé:")
    print(f"   Cellules remplies: {filled}")
    print(f"   Cellules sans mapping: {skipped}")
    print(f"   Fichier sauvegardé: {output_file}")
    
    return output_file


def main():
    """Pipeline complète"""
    
    print("\n" + "=" * 80)
    print(" SYSTEME COMPLET: Bilan GULFCAM - DSF")
    print("=" * 80)
    
    balance_file = "BALANCE AU 31 12 2024 VERSION DU 26 05 25.xlsx"
    
    try:
        # 1. Lire le bilan
        reader = BalanceReader(balance_file)
        reader.load()
        
        # 2. Extraire les totaux 2-chiffres
        print(f"\n📊 EXTRACTION COMPTES 2-CHIFFRES:")
        totals = reader.extract_two_digit_totals()
        
        print(f"\nComptes extraits ({len(totals)}):")
        total_amount = Decimal('0')
        for compte in sorted(totals.keys()):
            solde = totals[compte]
            print(f"   C{compte}: {solde:>18,.0f}")
            total_amount += solde
        
        print(f"\n   Total général: {total_amount:>18,.0f}")
        
        # 3. Remplir le DSF
        mapping = create_dsf_mapping()
        fill_dsf_template(totals, mapping)
        
        print("\n" + "=" * 80)
        print("✅ PIPELINE COMPLÈTE - DSF REMPLI AVEC SUCCÈS")
        print("=" * 80 + "\n")
        
    except Exception as e:
        print(f"\n❌ ERREUR: {e}\n")
        import traceback
        traceback.print_exc()


def fill_dsf_with_balance_values(balance_reader: BalanceReader, mapping: Dict, prefiller) -> Dict:
    """
    Remplir le DSF de manière intelligente et appropriée avec les valeurs du bilan.
    
    Cette fonction:
    1. Extrait les comptes 2-chiffres du bilan
    2. Analyze les zones de saisies du DSF
    3. Mappe chaque compte SYSCOHADA à sa zone de saisie appropriée
    4. Remplit uniquement les zones de saisies (pas les titres/labels)
    5. Retourne un rapport détaillé
    
    Args:
        balance_reader: BalanceReader chargé
        mapping: Dict du mapping SYSCOHADA → DSF
        prefiller: DSFPrefillerProduction pour écrire les cellules
        
    Returns:
        Dict avec résumé: {
            'total_filled': int,
            'total_mapped': int,
            'filled_details': List[(compte, label, sheet, cell, value)],
            'unmapped_accounts': List[(compte, label, value)]
        }
    """
    # Extraire les totaux de classe du bilan
    totals = balance_reader.extract_two_digit_totals()
    
    # Analyser le template DSF pour identifier les zones de saisies
    zones_by_sheet = _analyze_dsf_saisies(prefiller.wb)
    
    filled_details = []
    unmapped_accounts = []
    filled_count = 0
    mapped_count = len([c for c in mapping.keys() if c in totals])
    
    print(f"\n📊 REMPLISSAGE INTELLIGENT DSF DEPUIS BILAN")
    print(f"Comptes 2-chiffres trouvés: {len(totals)}")
    print(f"Comptes mappés disponibles: {mapped_count}")
    print("-" * 80)
    
    # Parcourir tous les comptes du bilan
    for compte in sorted(totals.keys()):
        value = totals[compte]
        
        # Trouver le libellé
        label = ""
        for c, lib, sol in balance_reader.read_all_accounts():
            if c == compte:
                label = lib
                break
        
        if compte in mapping:
            # Compte mappé - vérifier que c'est une zone de saisie
            map_info = mapping[compte]
            sheet = map_info["sheet"]
            cell = map_info["cell"]
            
            # Vérifier si c'est vraiment une zone de saisie
            is_saisie_zone = _is_saisie_cell(sheet, cell, zones_by_sheet)
            
            if is_saisie_zone:
                # Zone de saisie confirmée - remplir
                if prefiller.write_cell(sheet, cell, float(value)):
                    filled_details.append({
                        'compte': compte,
                        'label': label,
                        'sheet': sheet,
                        'cell': cell,
                        'value': value,
                        'zone_type': 'saisie'
                    })
                    filled_count += 1
                    print(f"✓ {compte:2s} → {sheet:20s} {cell:4s} = {value:>15,.0f} XAF  | {label[:30]}")
                else:
                    print(f"✗ {compte} : Erreur écriture cellule {cell}")
            else:
                # Pas une zone de saisie - avertissement
                print(f"⚠ {compte} : {cell} n'est pas une zone de saisie (titre/label)")
        else:
            # Compte non mappé
            unmapped_accounts.append({
                'compte': compte,
                'label': label,
                'value': value
            })
            print(f"⊘ {compte} : PAS DE MAPPING                           | {label[:30]}")
    
    print("-" * 80)
    print(f"Résumé: {filled_count} cellules remplies / {mapped_count} mappées")
    print(f"Comptes sans mapping: {len(unmapped_accounts)}")
    print("=" * 80)
    
    return {
        'total_filled': filled_count,
        'total_mapped': mapped_count,
        'filled_details': filled_details,
        'unmapped_accounts': unmapped_accounts
    }


def _analyze_dsf_saisies(wb) -> Dict:
    """Analyser le workbook pour identifier les zones de saisies"""
    zones_by_sheet = {}
    
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        saisies = set()
        
        # Parcourir les cellules pour trouver les zones vides/de saisie
        for row in ws.iter_rows(min_row=1, max_row=min(200, ws.max_row)):
            for cell in row:
                # Une zone de saisie = cellule vide ou avec nombre
                if cell.value is None:
                    # Cellule vide = zone de saisie
                    saisies.add(cell.coordinate)
                elif isinstance(cell.value, (int, float)):
                    # Nombre = zone de saisie avec donnée pré-remplie
                    saisies.add(cell.coordinate)
        
        zones_by_sheet[sheet_name] = {'saisies': saisies}
    
    return zones_by_sheet


def _is_saisie_cell(sheet: str, cell_ref: str, zones_by_sheet: Dict) -> bool:
    """Vérifier si une cellule est une zone de saisie"""
    if sheet not in zones_by_sheet:
        return False
    
    return cell_ref in zones_by_sheet[sheet].get('saisies', set())






