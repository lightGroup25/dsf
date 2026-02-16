# -*- coding: utf-8 -*-
"""
Module de transformation Balance de Comptes SYSCOHADA → Données DSF.
Architecture streaming pour éviter overflow mémoire.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Any, Generator
import re

# Lazy import pandas to avoid PySide6/six compatibility issues
pd = None

import openpyxl

from syscohada_db import SYSCOHADA_INDEX, get_account

logger = logging.getLogger(__name__)

def _ensure_pandas():
    """Ensure pandas is imported when needed."""
    global pd
    if pd is None:
        import pandas as _pd
        pd = _pd
    return pd


@dataclass
class BalanceRow:
    """Représente une ligne de balance de comptes."""
    compte_num: str
    compte_label: str
    solde_initial: float = 0.0
    debit: float = 0.0
    credit: float = 0.0
    solde_final: float = 0.0
    flux_tresorerie: float = 0.0  # Pour N-1 flux

    def to_dict(self) -> Dict[str, Any]:
        return {
            "compte": self.compte_num,
            "label": self.compte_label,
            "solde_initial": self.solde_initial,
            "debit": self.debit,
            "credit": self.credit,
            "solde_final": self.solde_final,
            "flux_tresorerie": self.flux_tresorerie,
        }


@dataclass
class BalanceLigneMapping:
    """Map un compte de balance à sa destination DSF."""
    compte_syscohada: str
    compte_label: str
    dsf_sheet: str
    dsf_cell: str  # Format "A1"
    dsf_field: str  # Ex: "ACTIF_COURANT"
    is_debit: bool = True  # True=débit, False=crédit
    multiply_by: float = 1.0  # Facteur multiplicatif
    condition: Optional[str] = None  # Regex ou condition


DSF_MAPPING: Dict[str, BalanceLigneMapping] = {
    # ===== CLASSE 1: CAPITAUX PROPRES ET RESSOURCES ASSIMILÉES =====
    
    # Sous-classe 10 - Capital et apporteurs
    "101": BalanceLigneMapping("101", "CAPITAL SOCIAL", "BILAN PAYSAGE", "D33", "EQUITY_CAPITAL", True),  # CA
    "102": BalanceLigneMapping("102", "APPORTEURS CAPITAL NON APPELÉ", "BILAN PAYSAGE", "D34", "EQUITY_UNCALLED", False),  # CB
    "104": BalanceLigneMapping("104", "PRIMES CAPITAL", "BILAN PAYSAGE", "D35", "EQUITY_PREMIUMS", True),  # CD
    "105": BalanceLigneMapping("105", "ÉCARTS DE RÉÉVALUATION", "BILAN PAYSAGE", "D36", "EQUITY_REVALUATION", True),  # CE
    
    # Sous-classe 11 - Réserves indisponibles  
    "110": BalanceLigneMapping("110", "RÉSERVES INDISPONIBLES", "BILAN PAYSAGE", "D37", "RESERVES_RESTRICTED", True),  # CF
    "111": BalanceLigneMapping("111", "RÉSERVE LÉGALE", "BILAN PAYSAGE", "D37", "RESERVES_RESTRICTED", True),
    "112": BalanceLigneMapping("112", "RÉSERVES STATUTAIRES", "BILAN PAYSAGE", "D37", "RESERVES_RESTRICTED", True),
    "118": BalanceLigneMapping("118", "AUTRES RÉSERVES INDISPONIBLES", "BILAN PAYSAGE", "D37", "RESERVES_RESTRICTED", True),
    
    # Sous-classe 12 - Réserves libres
    "120": BalanceLigneMapping("120", "RÉSERVES LIBRES", "BILAN PAYSAGE", "D38", "RESERVES_FREE", True),  # CG
    "121": BalanceLigneMapping("121", "RÉSULTATS AFFECTÉS", "BILAN PAYSAGE", "D38", "RESERVES_FREE", True),
    
    # Sous-classe 13 - Report à nouveau
    "130": BalanceLigneMapping("130", "REPORT À NOUVEAU", "BILAN PAYSAGE", "D39", "RETAINED_EARNINGS", True),  # CH
    "131": BalanceLigneMapping("131", "BÉNÉFICES REPORTÉS", "BILAN PAYSAGE", "D39", "RETAINED_EARNINGS", True),
    "132": BalanceLigneMapping("132", "PERTES REPORTÉES", "BILAN PAYSAGE", "D39", "RETAINED_EARNINGS", False),
    
    # Sous-classe 14 - Résultat net de l'exercice
    "140": BalanceLigneMapping("140", "RÉSULTAT NET EXERCICE", "BILAN PAYSAGE", "D40", "NET_RESULT", True),  # CJ
    
    # Sous-classe 15 - Subventions d'investissement
    "150": BalanceLigneMapping("150", "SUBVENTIONS INVESTISSEMENT", "BILAN PAYSAGE", "D41", "SUBSIDIES", True),  # CL
    
    # Sous-classe 16 - Provisions réglementées
    "160": BalanceLigneMapping("160", "PROVISIONS RÉGLEMENTÉES", "BILAN PAYSAGE", "D42", "PROVISIONS_REGULATED", True),  # CM
    
    # ===== CLASSE 2: IMMOBILISATIONS =====
    
    # Sous-classe 20 - Immobilisations incorporelles
    "201": BalanceLigneMapping("201", "FRAIS DÉVELOPPEMENT", "BILAN PAYSAGE", "D14", "INTANGIBLE_DEVELOPMENT", True),  # AE
    "202": BalanceLigneMapping("202", "BREVETS LICENCES SOFTWARE", "BILAN PAYSAGE", "D15", "INTANGIBLE_PATENTS", True),  # AF
    "203": BalanceLigneMapping("203", "FONDS COMMERCIAL", "BILAN PAYSAGE", "D16", "INTANGIBLE_GOODWILL", True),  # AG
    "204": BalanceLigneMapping("204", "AUTRES IMMOBILISATIONS INCORPORELLES", "BILAN PAYSAGE", "D17", "INTANGIBLE_OTHER", True),  # AH
    "2011": BalanceLigneMapping("2011", "FRAIS DE RECHERCHE", "BILAN PAYSAGE", "D14", "INTANGIBLE_DEVELOPMENT", True),
    "2012": BalanceLigneMapping("2012", "FRAIS FORMATION INFORMATIQUE", "BILAN PAYSAGE", "D14", "INTANGIBLE_DEVELOPMENT", True),
    "2013": BalanceLigneMapping("2013", "FRAIS PRÉ-EXPLOITATION", "BILAN PAYSAGE", "D14", "INTANGIBLE_DEVELOPMENT", True),
    
    # Sous-classe 21 - Immobilisations corporelles
    "210": BalanceLigneMapping("210", "TERRAINS", "BILAN PAYSAGE", "D18", "TANGIBLE_LAND", True),  # AJ
    "211": BalanceLigneMapping("211", "BÂTIMENTS", "BILAN PAYSAGE", "D19", "TANGIBLE_BUILDINGS", True),  # AK
    "212": BalanceLigneMapping("212", "AMÉNAGEMENTS INSTALLATIONS", "BILAN PAYSAGE", "D20", "TANGIBLE_FIXTURES", True),  # AL
    "213": BalanceLigneMapping("213", "MATÉRIEL MOBILIER", "BILAN PAYSAGE", "D21", "TANGIBLE_EQUIPMENT", True),  # AM
    "214": BalanceLigneMapping("214", "MATÉRIEL TRANSPORT", "BILAN PAYSAGE", "D22", "TANGIBLE_TRANSPORT", True),  # AN
    "215": BalanceLigneMapping("215", "ANIMAUX D'ÉLEVAGE", "BILAN PAYSAGE", "D18", "TANGIBLE_ANIMALS", True),
    "216": BalanceLigneMapping("216", "PLANTATIONS", "BILAN PAYSAGE", "D18", "TANGIBLE_PLANTATIONS", True),
    "217": BalanceLigneMapping("217", "SEMENCES REPRODUTRICES", "BILAN PAYSAGE", "D18", "TANGIBLE_SEEDS", True),
    "218": BalanceLigneMapping("218", "AUTRES ACTIFS BIOLOGIQUES", "BILAN PAYSAGE", "D18", "TANGIBLE_BIOLOGY", True),
    "219": BalanceLigneMapping("219", "ACTIFS EN COURS DE CONSTRUCTION", "BILAN PAYSAGE", "D18", "TANGIBLE_CONSTRUCTION", True),
    
    # Sous-classe 23 - Avances sur immobilisations
    "231": BalanceLigneMapping("231", "AVANCES IMMOBILISATIONS", "BILAN PAYSAGE", "D23", "ADVANCES_FIXED", True),  # AP
    
    # Sous-classe 24 - Immobilisations financières
    "241": BalanceLigneMapping("241", "TITRES PARTICIPATION", "BILAN PAYSAGE", "D24", "SECURITIES_PARTICIPATION", True),  # AR
    "248": BalanceLigneMapping("248", "AUTRES IMMOBILISATIONS FINANCIÈRES", "BILAN PAYSAGE", "D25", "SECURITIES_OTHER", True),  # AS
    
    # Amortissements (sous-classe 28-29)
    "281": BalanceLigneMapping("281", "AMORT IMMOBILISATIONS INCORPORELLES", "BILAN PAYSAGE", "E14", "AMORT_INTANGIBLE", False),
    "282": BalanceLigneMapping("282", "AMORT IMMOBILISATIONS CORPORELLES", "BILAN PAYSAGE", "E18", "AMORT_TANGIBLE", False),
    "2811": BalanceLigneMapping("2811", "AMORT FRAIS DÉVELOPPEMENT", "BILAN PAYSAGE", "E14", "AMORT_INTANGIBLE", False),
    "2812": BalanceLigneMapping("2812", "AMORT BREVETS LICENCES", "BILAN PAYSAGE", "E14", "AMORT_INTANGIBLE", False),
    "2821": BalanceLigneMapping("2821", "AMORT IMMEUBLES", "BILAN PAYSAGE", "E18", "AMORT_TANGIBLE", False),
    "2822": BalanceLigneMapping("2822", "AMORT AMÉNAGEMENTS", "BILAN PAYSAGE", "E18", "AMORT_TANGIBLE", False),
    "2823": BalanceLigneMapping("2823", "AMORT MATÉRIEL", "BILAN PAYSAGE", "E18", "AMORT_TANGIBLE", False),
    "2824": BalanceLigneMapping("2824", "AMORT MATÉRIEL TRANSPORT", "BILAN PAYSAGE", "E18", "AMORT_TANGIBLE", False),
    
    # ===== CLASSE 3: STOCKS =====
    
    "310": BalanceLigneMapping("310", "MATIÈRES PREMIÈRES", "BILAN PAYSAGE", "D26", "INVENTORY_RAW", True),  # BA
    "311": BalanceLigneMapping("311", "MATIÈRES CONSOMMABLES", "BILAN PAYSAGE", "D27", "INVENTORY_CONSUMED", True),  # BB
    "320": BalanceLigneMapping("320", "EN-COURS DE PRODUCTION", "BILAN PAYSAGE", "D28", "INVENTORY_WIP", True),  # BC
    "330": BalanceLigneMapping("330", "PRODUITS FINIS", "BILAN PAYSAGE", "D29", "INVENTORY_FINISHED", True),  # BD
    "340": BalanceLigneMapping("340", "MARCHANDISES", "BILAN PAYSAGE", "D30", "INVENTORY_GOODS", True),  # BE
    "350": BalanceLigneMapping("350", "STOCKS EN TRANSIT", "BILAN PAYSAGE", "D31", "INVENTORY_TRANSIT", True),  # BF
    "391": BalanceLigneMapping("391", "DÉPRÉCIATIONS STOCKS", "BILAN PAYSAGE", "D32", "INVENTORY_DEPRECIATION", False),
    
    # ===== CLASSE 4: TIERS / CRÉANCES / DETTES =====
    
    # Clients et créances
    "401": BalanceLigneMapping("401", "CLIENTS", "BILAN PAYSAGE", "D11", "RECEIVABLES_CLIENTS", True),  # CL
    "408": BalanceLigneMapping("408", "CLIENTS FACTURES À ÉTABLIR", "BILAN PAYSAGE", "D11", "RECEIVABLES_CLIENTS", True),
    "411": BalanceLigneMapping("411", "CLIENTS FACTURES REÇUES", "BILAN PAYSAGE", "D11", "RECEIVABLES_CLIENTS", True),
    "417": BalanceLigneMapping("417", "CLIENTS PRODUITS NON FACTURÉS", "BILAN PAYSAGE", "D11", "RECEIVABLES_CLIENTS", True),
    
    # Fournisseurs et dettes
    "501": BalanceLigneMapping("501", "FOURNISSEURS", "BILAN PAYSAGE", "K51", "PAYABLES_SUPPLIERS", False),  # DB
    "404": BalanceLigneMapping("404", "FOURNISSEURS FACTURES À RECEVOIR", "BILAN PAYSAGE", "K51", "PAYABLES_SUPPLIERS", False),
    "407": BalanceLigneMapping("407", "FOURNISSEURS FACTURES REÇUES", "BILAN PAYSAGE", "K51", "PAYABLES_SUPPLIERS", False),
    
    # Personnel et dettes sociales
    "421": BalanceLigneMapping("421", "PERSONNEL - SALAIRES DUES", "BILAN PAYSAGE", "K52", "PAYABLES_PERSONNEL", False),
    "422": BalanceLigneMapping("422", "PERSONNEL - CONGÉS PAYÉS", "BILAN PAYSAGE", "K52", "PAYABLES_PERSONNEL", False),
    
    # Organismes sociaux
    "431": BalanceLigneMapping("431", "ORGANISMES SOCIAUX", "BILAN PAYSAGE", "K53", "PAYABLES_SOCIAL", False),
    "4311": BalanceLigneMapping("4311", "CNPS COTISATIONS PATRONALES", "BILAN PAYSAGE", "K53", "PAYABLES_SOCIAL", False),
    
    # État - impôts
    "441": BalanceLigneMapping("441", "ÉTAT - IMPÔTS DIRECTS", "BILAN PAYSAGE", "K54", "PAYABLES_TAXES", False),
    "4411": BalanceLigneMapping("4411", "IMPÔT SOCIÉTÉ", "BILAN PAYSAGE", "K54", "PAYABLES_TAXES", False),
    "4412": BalanceLigneMapping("4412", "IMPÔT SYNTHÉTIQUE", "BILAN PAYSAGE", "K54", "PAYABLES_TAXES", False),
    
    # TVA
    "4451": BalanceLigneMapping("4451", "TVA FACTURÉE", "BILAN PAYSAGE", "K54", "PAYABLES_TAXES", True),
    "4452": BalanceLigneMapping("4452", "TVA RÉCUPÉRABLE", "BILAN PAYSAGE", "K54", "PAYABLES_TAXES", True),
    
    # Comptes courants associés
    "455": BalanceLigneMapping("455", "COMPTES COURANTS ASSOCIÉS", "BILAN PAYSAGE", "K55", "PAYABLES_ASSOCIATES", False),
    
    # Autres tiers
    "461": BalanceLigneMapping("461", "DÉBITEURS DIVERS", "BILAN PAYSAGE", "D12", "RECEIVABLES_OTHER", True),
    "471": BalanceLigneMapping("471", "CRÉDITEURS DIVERS", "BILAN PAYSAGE", "K55", "PAYABLES_OTHER", False),
    
    # ===== CLASSE 5: COMPTES DE TRÉSORERIE =====
    
    "500": BalanceLigneMapping("500", "TITRES DE PLACEMENT", "BILAN PAYSAGE", "D9", "SECURITIES", True),
    "501": BalanceLigneMapping("501", "ACTIONS", "BILAN PAYSAGE", "D9", "SECURITIES", True),
    "502": BalanceLigneMapping("502", "OBLIGATIONS", "BILAN PAYSAGE", "D9", "SECURITIES", True),
    "510": BalanceLigneMapping("510", "PRÊTS", "BILAN PAYSAGE", "D9", "LOANS", True),
    "520": BalanceLigneMapping("520", "COMPTES COURANTS BANCAIRES", "BILAN PAYSAGE", "D8", "BANK_ACCOUNTS", True),
    "521": BalanceLigneMapping("521", "COMPTES COURANTS PRINCIPAUX", "BILAN PAYSAGE", "D8", "BANK_ACCOUNTS", True),
    "530": BalanceLigneMapping("530", "CAISSE", "BILAN PAYSAGE", "D10", "CASH", True),
    
    # ===== CLASSE 6: CHARGES (RÉSULTATS) =====
    
    "600": BalanceLigneMapping("600", "ACHATS DE MARCHANDISES", "COMPTE DE RESULTAT", "D7", "PURCHASES_GOODS", False),
    "601": BalanceLigneMapping("601", "ACHATS MATIÈRES PREMIÈRES", "COMPTE DE RESULTAT", "D8", "PURCHASES_RAW", False),
    "606": BalanceLigneMapping("606", "ACHATS FOURNITURES", "COMPTE DE RESULTAT", "D9", "PURCHASES_SUPPLIES", False),
    
    "610": BalanceLigneMapping("610", "VARIATION STOCKS MARCHANDISES", "COMPTE DE RESULTAT", "D10", "INVENTORY_VARIATION", False),
    "611": BalanceLigneMapping("611", "VARIATION STOCKS MATIÈRES PREMIÈRES", "COMPTE DE RESULTAT", "D11", "INVENTORY_VARIATION_RAW", False),
    
    "620": BalanceLigneMapping("620", "TRANSPORT MARCHANDISES", "COMPTE DE RESULTAT", "D15", "TRANSPORT_GOODS", False),
    "621": BalanceLigneMapping("621", "CARBURANT", "COMPTE DE RESULTAT", "D15", "FUEL", False),
    "622": BalanceLigneMapping("622", "ENTRETIEN ÉQUIPEMENTS", "COMPTE DE RESULTAT", "D15", "MAINTENANCE", False),
    "623": BalanceLigneMapping("623", "ENTRETIEN IMMEUBLES", "COMPTE DE RESULTAT", "D15", "MAINTENANCE_BUILDINGS", False),
    "624": BalanceLigneMapping("624", "LOYERS", "COMPTE DE RESULTAT", "D16", "RENT", False),
    "625": BalanceLigneMapping("625", "ASSURANCES", "COMPTE DE RESULTAT", "D17", "INSURANCE", False),
    "626": BalanceLigneMapping("626", "FRAIS VOYAGE", "COMPTE DE RESULTAT", "D18", "TRAVEL", False),
    "627": BalanceLigneMapping("627", "TÉLÉCOMMUNICATIONS", "COMPTE DE RESULTAT", "D19", "TELECOM", False),
    "628": BalanceLigneMapping("628", "AUTRES SERVICES", "COMPTE DE RESULTAT", "D20", "SERVICES_OTHER", False),
    
    "630": BalanceLigneMapping("630", "IMPÔTS ET TAXES", "COMPTE DE RESULTAT", "D21", "TAXES_DUTIES", False),
    "631": BalanceLigneMapping("631", "IMPÔTS PATRIMOINE", "COMPTE DE RESULTAT", "D21", "TAXES_PROPERTY", False),
    "632": BalanceLigneMapping("632", "PATENTES LICENCES", "COMPTE DE RESULTAT", "D21", "TAXES_LICENSES", False),
    
    "640": BalanceLigneMapping("640", "SALAIRES ET TRAITEMENTS", "COMPTE DE RESULTAT", "D25", "SALARIES", False),
    "641": BalanceLigneMapping("641", "COTISATIONS PATRONALES CNPS", "COMPTE DE RESULTAT", "D26", "SOCIAL_CONTRIBUTIONS", False),
    "642": BalanceLigneMapping("642", "COTISATIONS MUTUELLES", "COMPTE DE RESULTAT", "D26", "MUTUAL_CONTRIBUTIONS", False),
    "643": BalanceLigneMapping("643", "FRAIS MÉDICAUX", "COMPTE DE RESULTAT", "D27", "MEDICAL_EXPENSES", False),
    "644": BalanceLigneMapping("644", "COTISATIONS ACCIDENTS TRAVAIL", "COMPTE DE RESULTAT", "D26", "ACCIDENT_INSURANCE", False),
    
    "650": BalanceLigneMapping("650", "AMORTISSEMENTS", "COMPTE DE RESULTAT", "D28", "DEPRECIATION", False),
    "6511": BalanceLigneMapping("6511", "DÉPRÉCIATIONS IMMOBILISATIONS", "COMPTE DE RESULTAT", "D29", "IMPAIRMENT_FIXED", False),
    "6512": BalanceLigneMapping("6512", "DÉPRÉCIATIONS STOCKS", "COMPTE DE RESULTAT", "D30", "IMPAIRMENT_INVENTORY", False),
    "6513": BalanceLigneMapping("6513", "DÉPRÉCIATIONS CRÉANCES", "COMPTE DE RESULTAT", "D31", "IMPAIRMENT_RECEIVABLES", False),
    
    "660": BalanceLigneMapping("660", "PROVISION EXERCICE", "COMPTE DE RESULTAT", "D32", "PROVISIONS", False),
    "6611": BalanceLigneMapping("6611", "PROVISIONS LITIGES", "COMPTE DE RESULTAT", "D32", "PROVISIONS_LITIGATION", False),
    
    "670": BalanceLigneMapping("670", "CHARGES D'INTÉRÊTS", "COMPTE DE RESULTAT", "D35", "INTEREST_EXPENSE", False),
    "671": BalanceLigneMapping("671", "AUTRES CHARGES FINANCIÈRES", "COMPTE DE RESULTAT", "D36", "FINANCIAL_CHARGES_OTHER", False),
    "672": BalanceLigneMapping("672", "PERTES CHANGES", "COMPTE DE RESULTAT", "D37", "FX_LOSSES", False),
    
    "680": BalanceLigneMapping("680", "CHARGES EXCEPTIONNELLES", "COMPTE DE RESULTAT", "D40", "EXCEPTIONAL_CHARGES", False),
    "681": BalanceLigneMapping("681", "PERTES ÉLÉMENTS D'ACTIF", "COMPTE DE RESULTAT", "D40", "ASSET_LOSSES", False),
    
    # ===== CLASSE 7: PRODUITS (RÉSULTATS) =====
    
    "700": BalanceLigneMapping("700", "VENTES DE MARCHANDISES", "COMPTE DE RESULTAT", "D5", "SALES_GOODS", True),
    "701": BalanceLigneMapping("701", "VENTES DE PRODUITS FINIS", "COMPTE DE RESULTAT", "D6", "SALES_PRODUCTS", True),
    "702": BalanceLigneMapping("702", "TRAVAUX EN PROGRÈS", "COMPTE DE RESULTAT", "D6", "WORK_PROGRESS", True),
    "703": BalanceLigneMapping("703", "PRESTATIONS DE SERVICES", "COMPTE DE RESULTAT", "D6", "SERVICES_REVENUE", True),
    "706": BalanceLigneMapping("706", "AUTRES REVENUS D'EXPLOITATION", "COMPTE DE RESULTAT", "D6", "REVENUE_OTHER", True),
    "708": BalanceLigneMapping("708", "PÉNALITÉS FACTURÉES", "COMPTE DE RESULTAT", "D6", "PENALTIES", True),
    
    "710": BalanceLigneMapping("710", "REVENUS LOYERS ET FERMAGES", "COMPTE DE RESULTAT", "D12", "RENTAL_INCOME", True),
    "711": BalanceLigneMapping("711", "REVENUS LOCATIONS IMMOBILISATIONS", "COMPTE DE RESULTAT", "D12", "LEASING_INCOME", True),
    
    "720": BalanceLigneMapping("720", "PRODUITS PLACEMENTS", "COMPTE DE RESULTAT", "D13", "INVESTMENT_INCOME", True),
    "721": BalanceLigneMapping("721", "INTÉRÊTS PRÊTS", "COMPTE DE RESULTAT", "D14", "INTEREST_INCOME", True),
    
    "740": BalanceLigneMapping("740", "REPRISES AMORTISSEMENTS", "COMPTE DE RESULTAT", "D33", "DEPRECIATION_REVERSAL", True),
    "7411": BalanceLigneMapping("7411", "REPRISES PROVISIONS", "COMPTE DE RESULTAT", "D34", "PROVISIONS_REVERSAL", True),
    
    "750": BalanceLigneMapping("750", "GAINS CHANGES", "COMPTE DE RESULTAT", "D38", "FX_GAINS", True),
    
    "760": BalanceLigneMapping("760", "AUTRES PRODUITS FINANCIERS", "COMPTE DE RESULTAT", "D39", "FINANCIAL_PRODUCTS_OTHER", True),
    "770": BalanceLigneMapping("770", "PRODUITS EXCEPTIONNELS", "COMPTE DE RESULTAT", "D41", "EXCEPTIONAL_INCOME", True),
    
    # ===== CLASSE 8: COMPTES DE RÉSULTATS ANALYTIQUES (si applicable) =====
    
    "801": BalanceLigneMapping("801", "FRAIS DE PRODUCTION CENTRE 1", "COMPTE DE RESULTAT", "D42", "COST_CENTER_1", False),
    "802": BalanceLigneMapping("802", "FRAIS DE PRODUCTION CENTRE 2", "COMPTE DE RESULTAT", "D43", "COST_CENTER_2", False),
    
    # ===== CLASSE 9: COMPTES ENGAGEMENTS HORS BILAN =====
    
    "900": BalanceLigneMapping("900", "ENGAGEMENTS REÇUS", "TABLEAU DES FLUX DE TRESORERIE", "D5", "COMMITMENTS_RECEIVED", False),
    "901": BalanceLigneMapping("901", "ENGAGEMENTS DONNÉS", "TABLEAU DES FLUX DE TRESORERIE", "D6", "COMMITMENTS_GIVEN", False),
}


class BalanceStreamParser:
    """Parse une balance de comptes (fichier Excel) en stream."""

    def __init__(
        self,
        balance_file: Path,
        chunk_size: int = 500,
        column_overrides: Optional[Dict[str, object]] = None,
    ):
        """
        Args:
            balance_file: Chemin du fichier de balance
            chunk_size: Nombre de lignes à charger en mémoire à la fois
        """
        self.balance_file = Path(balance_file)
        self.chunk_size = chunk_size
        self.wb = None
        self.ws = None
        self.header_row = 1
        self.column_overrides = column_overrides or {}
        self._col_map: Dict[str, Any] = {}

    def load_balance_structure(self) -> Dict[str, int]:
        """
        Detect les colonnes de la balance (flexible selon année).
        
        Retourne un dict mapping: {
            'compte': col_index,
            'label': col_index,
            'solde_initial': col_index,
            'debit': col_index,
            'credit': col_index,
            'solde_final': col_index,
            'flux_n1': col_index (optional)
        }
        """
        self.wb = openpyxl.load_workbook(self.balance_file, data_only=True)
        self.ws = self.wb.active

        col_map = {}
        
        # If column overrides are provided, use them directly (special case for GULFCAM format)
        if self.column_overrides:
            col_map = self.column_overrides.copy()
            self.header_row = 8  # Force header row 8 for GULFCAM format (row 8 has "Numéro de compte")
            logger.info(f"Using column overrides (GULFCAM format): {col_map}")
            self._col_map = col_map
            return col_map
        
        # Otherwise, auto-detect header row
        header_row = self._detect_header_row()
        self.header_row = header_row

        for col_idx, cell in enumerate(self.ws.iter_cols(min_row=header_row, max_row=header_row), 1):
            header_text = str(cell[0].value or "").upper().strip()

            if "COMPTE" in header_text or "NUM" in header_text:
                col_map["compte"] = col_idx
            elif "LIB" in header_text or "INTITULE" in header_text or "LABEL" in header_text:
                col_map["label"] = col_idx
            elif "SOLDE" in header_text and "INITIAL" in header_text:
                col_map["solde_initial"] = col_idx
            elif "DÉBIT" in header_text or "DEBIT" in header_text:
                col_map["debit"] = col_idx
            elif "CRÉDIT" in header_text or "CREDIT" in header_text:
                col_map["credit"] = col_idx
            elif "SOLDE" in header_text and "FINAL" in header_text:
                col_map["solde_final"] = col_idx
            elif "FLUX" in header_text or "N-1" in header_text:
                col_map["flux_n1"] = col_idx

        self._apply_overrides(col_map)
        self._normalize_column_lists(col_map)
        self._col_map = col_map
        logger.info(f"Colonnes détectées (ligne {header_row}): {col_map}")
        return col_map

    def iter_balance_rows(self, col_map: Dict[str, int]) -> Generator[BalanceRow, None, None]:
        """
        Itère sur les lignes de balance en stream (yield).
        Évite de charger tout en mémoire.
        """
        if not self.ws:
            self.load_balance_structure()

        start_row = (self.header_row or 1) + 1
        compte_col = col_map.get("compte")
        label_col = col_map.get("label", compte_col)
        solde_initial_col = col_map.get("solde_initial")
        flux_col = col_map.get("flux_n1")
        debit_columns = col_map.get("debit_columns")
        credit_columns = col_map.get("credit_columns")
        balance_columns = col_map.get("balance_columns")
        single_debit_col = col_map.get("debit")
        single_credit_col = col_map.get("credit")
        single_balance_col = col_map.get("solde_final")

        for row_idx, row in enumerate(self.ws.iter_rows(min_row=start_row, values_only=True), start=start_row):
            try:
                compte = str(row[compte_col - 1] or "").strip()
                if not compte or not compte[0].isdigit():
                    continue  # Skip non-compte rows

                label_value = row[label_col - 1] if label_col else None
                label = str(label_value or "").strip() or compte

                debit_value = (
                    self._first_numeric(row, debit_columns)
                    if debit_columns
                    else self._value_from_column(row, single_debit_col)
                )
                credit_value = (
                    self._first_numeric(row, credit_columns)
                    if credit_columns
                    else self._value_from_column(row, single_credit_col)
                )
                balance_value = None
                if balance_columns:
                    balance_value = self._first_numeric(row, balance_columns)
                elif single_balance_col:
                    balance_value = self._value_from_column(row, single_balance_col)

                solde_final = debit_value - credit_value
                if solde_final == 0 and balance_value not in (None, 0):
                    solde_final = balance_value

                balance_row = BalanceRow(
                    compte_num=compte,
                    compte_label=label,
                    solde_initial=self._value_from_column(row, solde_initial_col),
                    debit=debit_value,
                    credit=credit_value,
                    solde_final=solde_final,
                    flux_tresorerie=self._value_from_column(row, flux_col),
                )
                yield balance_row
                
            except (ValueError, IndexError, TypeError) as e:
                logger.warning(f"Erreur parsage ligne {row_idx}: {e}")
                continue

    def close(self):
        if self.wb:
            self.wb.close()
            self.wb = None

    def _apply_overrides(self, col_map: Dict[str, Any]) -> None:
        if not self.column_overrides:
            return
        for key, value in self.column_overrides.items():
            col_map[key] = value

    def _normalize_column_lists(self, col_map: Dict[str, Any]) -> None:
        for key in ("debit_columns", "credit_columns", "balance_columns"):
            if key in col_map:
                col_map[key] = self._to_int_list(col_map[key])

    @staticmethod
    def _to_int_list(value: Any) -> List[int]:
        if value is None:
            return []
        if isinstance(value, int):
            return [value]
        if isinstance(value, (list, tuple, set)):
            return [int(v) for v in value if v]
        return [int(value)]

    @staticmethod
    def _value_from_column(row: Tuple[Any, ...], col_idx: Optional[int]) -> float:
        if not col_idx:
            return 0.0
        try:
            value = row[col_idx - 1]
        except IndexError:
            return 0.0
        try:
            return float(value or 0)
        except (TypeError, ValueError):
            return 0.0

    @staticmethod
    def _first_numeric(row: Tuple[Any, ...], columns: Optional[List[int]]) -> float:
        if not columns:
            return 0.0
        for col_idx in columns:
            try:
                value = row[col_idx - 1]
            except IndexError:
                continue
            if value in (None, ""):
                continue
            try:
                number = float(value)
            except (TypeError, ValueError):
                continue
            if number != 0:
                return number
        return 0.0

    def _detect_header_row(self) -> int:
        if not self.ws:
            return 1
        max_scan = min(self.ws.max_row, 50)
        for row_idx, row in enumerate(self.ws.iter_rows(min_row=1, max_row=max_scan), start=1):
            for cell in row:
                value = str(cell.value or "").upper()
                if any(keyword in value for keyword in ("COMPTE", "N°", "NUM", "LIB", "INTITULE")):
                    return row_idx
        return 1


class BalanceTransformer:
    """Transforme une balance en entrées DSF."""

    @staticmethod
    def match_account_to_dsf(compte_num: str) -> Optional[BalanceLigneMapping]:
        """Trouve le mapping DSF pour un compte SYSCOHADA."""
        # Exact match
        if compte_num in DSF_MAPPING:
            return DSF_MAPPING[compte_num]

        # Prefix match (ex: 10 pour 101, 102, etc.)
        for prefix in [compte_num[:3], compte_num[:2], compte_num[:1]]:
            for mapped_account, mapping in DSF_MAPPING.items():
                if mapped_account.startswith(prefix):
                    logger.debug(f"Prefix match: {compte_num} => {mapped_account}")
                    return mapping

        return None

    @staticmethod
    def apply_mapping(balance_row: BalanceRow) -> Optional[Dict[str, Any]]:
        """Transforme une ligne de balance en entrée DSF."""
        mapping = BalanceTransformer.match_account_to_dsf(balance_row.compte_num)
        if not mapping:
            logger.debug(f"Pas de mapping pour {balance_row.compte_num}")
            return None

        # Sélectionner la valeur appropriée
        if mapping.is_debit:
            value = balance_row.solde_final  # Débit: prendre solde final
        else:
            value = -balance_row.solde_final  # Crédit: négatif

        return {
            "compte_syscohada": balance_row.compte_num,
            "compte_label": balance_row.compte_label,
            "dsf_sheet": mapping.dsf_sheet,
            "dsf_cell": mapping.dsf_cell,
            "dsf_field": mapping.dsf_field,
            "value": value * mapping.multiply_by,
            "original_value": balance_row.solde_final,
        }


__all__ = [
    "BalanceRow",
    "BalanceLigneMapping",
    "BalanceStreamParser",
    "BalanceTransformer",
    "DSF_MAPPING",
]


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)

    balance_file = Path("BALANCE AU 31 12 2024 VERSION DU 26 05 25.xlsx")
    parser = BalanceStreamParser(balance_file)

    col_map = parser.load_balance_structure()
    print(f"Structure balance: {col_map}")

    count = 0
    for balance_row in parser.iter_balance_rows(col_map):
        dsf_entry = BalanceTransformer.apply_mapping(balance_row)
        if dsf_entry:
            print(f"Mapped: {balance_row.compte_num} => {dsf_entry['dsf_cell']}")
            count += 1
        if count >= 10:
            break

    parser.close()
    print(f"Total lignes avec mapping: {count}")
