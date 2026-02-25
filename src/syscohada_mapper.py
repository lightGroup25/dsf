#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Double Mapping SYSCOHADA : (DSF → SYSCOHADA) + (SYSCOHADA → Balance)
Permet le remplissage rapide de NOTE 20 et 34 sans fuzzy matching
"""

from dataclasses import dataclass
from typing import Dict, List, Optional, Set, Tuple
from decimal import Decimal

# ============================================================================
# MAPPING NOTE 20 : DSF STRUCTURE → SYSCOHADA CLASSE 2 (IMMOBILISATIONS)
# ============================================================================

@dataclass
class Note20Section:
    """Section de NOTE 20 (immobilisations par catégorie)"""
    section_name: str           # "Incorporelles", "Corporelles", "Financières"
    syscohada_class: str        # "201", "202", "203"
    dsf_rows: List[int]         # Lignes concernées dans NOTE 20
    account_prefixes: List[str] # Comptes SYSCOHADA "201xxx", "202xxx", etc.
    
NOTE20_MAPPINGS: List[Note20Section] = [
    Note20Section(
        section_name="Immobilisations Incorporelles",
        syscohada_class="201",
        dsf_rows=[],  # À extraire du template
        account_prefixes=["201", "202", "203"]
    ),
    Note20Section(
        section_name="Immobilisations Corporelles",
        syscohada_class="203",
        dsf_rows=[],
        account_prefixes=["203", "204", "205"]
    ),
    Note20Section(
        section_name="Immobilisations Financières",
        syscohada_class="26",
        dsf_rows=[],
        account_prefixes=["261", "262", "263", "264", "265", "266"]
    ),
]

# ============================================================================
# MAPPING NOTE 34 : DSF STRUCTURE → RATIOS FINANCIERS
# ============================================================================

@dataclass
class Note34Indicator:
    """Indicateur financier de NOTE 34"""
    indicator_name: str         # "Chiffre d'affaires", "Marge commerciale"
    syscohada_formula: str      # Comment le calculer depuis SYSCOHADA
    required_accounts: Set[str] # Comptes nécessaires (ex: {"701", "702", "801"})

NOTE34_MAPPINGS: Dict[str, Note34Indicator] = {
    "CHIFFRE_AFFAIRES": Note34Indicator(
        indicator_name="Chiffre d'affaires",
        syscohada_formula="SUM(701, 702, 703, 704, 705, 706)",
        required_accounts={"701", "702", "703", "704", "705", "706"}
    ),
    "MARGE_COMMERCIALE": Note34Indicator(
        indicator_name="Marge commerciale",
        syscohada_formula="CA - COGS",
        required_accounts={"601", "602", "603", "701", "702"}
    ),
    "VALEUR_AJOUTEE": Note34Indicator(
        indicator_name="Valeur ajoutée",
        syscohada_formula="Marge - (601+602+603+604)",
        required_accounts={"601", "602", "603", "604"}
    ),
}

# ============================================================================
# MAPPING BIDIRECTIONNEL : Comptes SYSCOHADA ↔ Balance Input
# ============================================================================

@dataclass
class SyscohadadAccountMapping:
    """Mappe un compte SYSCOHADA à ses colonnes dans la balance"""
    account_code: str           # "201100", "202000", etc.
    account_label: str
    debit_col_n: int            # Colonne pour débit clôture (N)
    credit_col_n: int           # Colonne pour crédit clôture (N)
    debit_col_n1: int           # Colonne pour débit ouverture (N-1)
    credit_col_n1: int          # Colonne pour crédit ouverture (N-1)

class SyscohadadMapper:
    """Gère le double mapping SYSCOHADA"""
    
    def __init__(self):
        # Colonnes standard dans la balance GULFCAM
        self.debit_col_n = 24      # Débit clôture
        self.credit_col_n = 27     # Crédit clôture
        self.debit_col_n1 = 11     # Débit ouverture
        self.credit_col_n1 = 14    # Crédit ouverture
        
        # Cache des comptes par classe
        self.accounts_by_class: Dict[str, List[str]] = {}
        
    def get_accounts_for_class(self, classe: str, accounts: Dict[str, any]) -> List[str]:
        """Retourne tous les comptes d'une classe SYSCOHADA"""
        if classe not in self.accounts_by_class:
            matching = []
            for compte, row in accounts.items():
                if str(compte).startswith(classe):
                    matching.append(compte)
            self.accounts_by_class[classe] = matching
        return self.accounts_by_class[classe]
    
    def get_accounts_by_prefix_range(
        self, 
        start_prefix: str,
        end_prefix: str,
        accounts: Dict[str, any]
    ) -> List[str]:
        """Retourne comptes dans une plage (ex: 201000 à 209999)"""
        start_int = int(start_prefix)
        end_int = int(end_prefix)
        matching = []
        for compte, row in accounts.items():
            try:
                compte_int = int(compte)
                if start_int <= compte_int <= end_int:
                    matching.append(compte)
            except (ValueError, TypeError):
                pass
        return matching
    
    def map_note20_section_to_accounts(
        self,
        section: Note20Section,
        accounts: Dict[str, any]
    ) -> Dict[str, Tuple[Decimal, Decimal]]:  # {compte: (solde_n, solde_n1)}
        """Mappe une section NOTE 20 aux comptes correspondants"""
        result = {}
        
        # Pour chaque préfixe de classe, récupère tous les comptes
        for prefix in section.account_prefixes:
            matching = self.get_accounts_for_class(prefix, accounts)
            for compte in matching:
                row = accounts[compte]
                # Utilise solde_final qui est pré-calculé
                result[compte] = (row.solde_final, getattr(row, 'solde_n1', Decimal(0)))
        
        return result
    
    def map_note34_indicator(
        self,
        indicator: Note34Indicator,
        accounts: Dict[str, any]
    ) -> Tuple[Decimal, Decimal]:  # (valeur_n, valeur_n1)
        """Calcule une valeur NOTE 34 à partir des comptes SYSCOHADA"""
        value_n = Decimal(0)
        value_n1 = Decimal(0)
        
        for account_prefix in indicator.required_accounts:
            matching = self.get_accounts_for_class(account_prefix, accounts)
            for compte in matching:
                row = accounts[compte]
                value_n += Decimal(str(row.solde_final)) if row.solde_final else Decimal(0)
                value_n1 += Decimal(str(getattr(row, 'solde_n1', 0))) if hasattr(row, 'solde_n1') else Decimal(0)
        
        return (value_n, value_n1)


# ============================================================================
# OPTIMISATIONS SPÉCIFIQUES NOTE 20
# ============================================================================

class Note20FastFiller:
    """Remplis NOTE 20 de manière structurée sans fuzzy matching"""
    
    def __init__(self, mapper: SyscohadadMapper):
        self.mapper = mapper
        
    def fill_note20_section(
        self,
        ws,                    # Worksheet
        section: Note20Section,
        accounts: Dict[str, any],
        start_row: int,
        end_row: int
    ) -> int:
        """Remplit une section de NOTE 20 de manière structurée"""
        assignments_count = 0
        
        # Récupère tous les comptes pour cette section
        section_accounts = self.mapper.map_note20_section_to_accounts(section, accounts)
        
        # Pour chaque ligne de la section
        account_list = sorted(section_accounts.keys())
        row_idx = 0
        
        for excel_row in range(start_row, end_row + 1):
            if row_idx >= len(account_list):
                break
            
            compte = account_list[row_idx]
            solde_n, solde_n1 = section_accounts[compte]
            
            # Remplis les colonnes : compte (col A), libellé (col B), solde N-1 (col C), solde N (col D)
            ws.cell(excel_row, 1).value = compte
            row = accounts[compte]
            ws.cell(excel_row, 2).value = row.label
            
            # Formats spécifiques pour NOTE 20
            ws.cell(excel_row, 3).value = float(solde_n1) if solde_n1 else 0.0
            ws.cell(excel_row, 4).value = float(solde_n) if solde_n else 0.0
            
            row_idx += 1
            assignments_count += 1
        
        return assignments_count


if __name__ == "__main__":
    mapper = SyscohadadMapper()
    print("[OK] Mapper SYSCOHADA initialized")
    print(f"  Debit N:   col {mapper.debit_col_n}")
    print(f"  Credit N:  col {mapper.credit_col_n}")
    print(f"  Debit N-1: col {mapper.debit_col_n1}")
    print(f"  Credit N-1: col {mapper.credit_col_n1}")
