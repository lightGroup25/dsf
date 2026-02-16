#!/usr/bin/env python3
"""Automatic DSF rule generator based on the double mapping JSON.

The goal is to convert the SYSCOHADA → Intermediate → DSF mapping into a
rule file consumable by the dynamic rule engine. We rely on:

* dsf_final_mapping.json: provides the intermediate categories, their
  prefixes and the coverage of the 1182 SYSCOHADA accounts.
* data/dsf_inventory.json: describes every writable cell in the DSF
  template so we can validate that each generated rule points to an
  actual column segment.

The output is a YAML/JSON file that follows the structure consumed by
`dsf_rule_config.py`/`dsf_rule_engine.py`.
"""
from __future__ import annotations

import argparse
import json
import logging
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

from dsf_inventory import DSFInventory, InventoryField
from dsf_rule_config import (
    AccountFilter,
    DSFRule,
    DSFRuleSet,
    DSFRuleTarget,
    save_rule_set,
)

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Metadata describing how to transform abstract categories into concrete
# rule targets. This is intentionally coarse-grained: we let the rule engine
# pick the best cell inside the provided column window based on label
# similarity.
# ---------------------------------------------------------------------------
SHEET_ALIASES: Dict[str, str] = {
    "BILAN": "BILAN PAYSAGE",
    "COMPTE_RES": "COMPTE DE RESULTAT",
    "R1": "Fiche R1",
    "R2": "Fiche R2",
    "R3": "Fiche R3",
    "NOTE13": "NOTE 13 ",
    "TABLEAU DES FLUX DE TRESORERIE": "TABLEAU DES FLUX DE TRESORERIE",
    "NOTE 3A": "NOTE 3A",
    "NOTE 3C": "NOTE  3C",
    "NOTE_3E": "NOTE 3E ",
    "NOTE_4": "NOTE 4 ",
    "NOTE_5": "NOTE 5 ",
    "NOTE_6": "NOTE 6 ",
    "NOTE_7": "NOTE 7 ",
    "NOTE_8": "NOTE 8 ",
    "NOTE_9": "NOTE 9 ",
    "NOTE_10": "NOTE 10 ",
    "NOTE_11": "NOTE 11 ",
    "NOTE_12": "NOTE 12 ",
    "NOTE_13": "NOTE 13 ",  # Assuming "NOTE13" alias in structure or "NOTE 13 "
    "NOTE_14": "NOTE 14   ",
    "NOTE_15": "NOTE 15 ", # Verify? "NOTE 15A  " exists, "NOTE 15" ? I don't see "NOTE 15" strictly.
    "NOTE_15A": "NOTE 15A  ",
    "NOTE_16": "NOTE 16 ", # Not in list? "NOTE 16A ", "NOTE 16B", "NOTE 18"
    # Wait, where is Note 16? It must be 16A and 16B combined or just 16A.
    # The output shows 'NOTE 16A ', 'NOTE 16B', 'NOTE 16B BIS', 'NOTE 16C'.
    # So I should map "NOTE_16" -> "NOTE 16A " or split.
    "NOTE_16A": "NOTE 16A ",
    "NOTE_16B": "NOTE 16B",
    "NOTE_17": "NOTE 17  ",
    "NOTE_18": "NOTE 18",
    "NOTE_19": "NOTE 19 ",
    "NOTE_20": "NOTE 20 ",
    "NOTE_21": "NOTE 21",
    "NOTE_22": "NOTE 22",
    "NOTE_23": "NOTE 23 ",
    "NOTE_24": "NOTE 24 ",
    "NOTE_25": "NOTE 25 ",
    "NOTE_26": "NOTE 26 ",
    "NOTE_27": "NOTE 27 ",
    "NOTE_27A": "NOTE 27A",
    "NOTE_28": "NOTE 28",
    "NOTE_34": "NOTE 34",
    "PAGE_GARDE": "PAGE DE GARDE",
}

SIGN_BY_CLASS = {
    "1": "credit",
    "2": "debit",
    "3": "debit",
    "4": "credit",
    "5": "credit",
    "6": "debit",
    "7": "credit",
    # Classes 8 and 9 carry auxiliary information; we keep them sign-agnostic.
}

# Windows (row ranges + column) for each logical group.
CATEGORY_WINDOWS = {
    "bilan_asset_fixed": {"column": "D", "row_start": 12, "row_end": 27,
                           "section": "BILAN_ACTIF", "buckets": ["ACTIF", "IMMOBILISATIONS"]},
    "bilan_asset_current": {"column": "D", "row_start": 28, "row_end": 40,
                             "section": "BILAN_ACTIF", "buckets": ["ACTIF", "CIRCULANT"]},
    "bilan_equity": {"column": "K", "row_start": 12, "row_end": 22,
                      "section": "BILAN_PASSIF", "buckets": ["PASSIF", "CAPITAUX_PROPRES"]},
    "bilan_liability": {"column": "K", "row_start": 23, "row_end": 45,
                         "section": "BILAN_PASSIF", "buckets": ["PASSIF", "DETTES"]},
    "cr_revenue": {"column": "E", "row_start": 11, "row_end": 24,
                    "section": "COMPTE_RESULTAT_PRODUITS", "buckets": ["CR_PRODUITS"]},
    "cr_charge": {"column": "E", "row_start": 24, "row_end": 60,
                   "section": "COMPTE_RESULTAT_CHARGES", "buckets": ["CR_CHARGES"]},
    
    # Extensions for Special/Control/Analytic categories
    "special_provisions": {"column": "D", "row_start": 33, "row_end": 35,
                           "section": "BILAN_PASSIF", "buckets": ["PASSIF", "SPECIAL_PROVISIONS"]},
    "special_other": {"column": "D", "row_start": 40, "row_end": 45, # Example fallback
                       "section": "BILAN_PASSIF", "buckets": ["PASSIF", "SPECIAL_OTHER"]},
                       
    # NOTES
    "note_3a_fixed_assets": {"column": "D", "column_keyword": "BRUTE", "row_start": 15, "row_end": 30,
                             "section": "NOTE_3A", "buckets": ["NOTE_3A", "BRUT_FIN"]},
    "note_3c_amortization": {"column": "J", "column_keyword": "AMORTISSEMENT", "row_start": 15, "row_end": 30,
                             "section": "NOTE_3C", "buckets": ["NOTE_3C", "AMORT_FIN"]},

    # NOUVELLES NOTES (Standard SYSCOHADA)
    "note_4_financial_assets": {"column": "D", "column_keyword": "BRUT", "row_start": 10, "row_end": 40,
                      "section": "NOTE_4", "buckets": ["NOTE_4", "FINANCIAL"]},
    
    "note_5_receivables_hao": {"column": "D", "column_keyword": "BRUT", "row_start": 10, "row_end": 40,
                           "section": "NOTE_5", "buckets": ["NOTE_5", "HAO"]},
                           
    "note_6_inventory": {"column": "D", "column_keyword": "BRUT", "row_start": 10, "row_end": 40,
                      "section": "NOTE_6", "buckets": ["NOTE_6", "STOCKS"]},
                      
    "note_7_receivables": {"column": "D", "column_keyword": "BRUT", "row_start": 10, "row_end": 60,
                           "section": "NOTE_7", "buckets": ["NOTE_7", "CLIENTS"]},
                           
    "note_8_other_receivables": {"column": "D", "column_keyword": "BRUT", "row_start": 10, "row_end": 40,
                                 "section": "NOTE_8", "buckets": ["NOTE_8", "OTHER_RECEIVABLES"]},
                                 
    "note_9_securities": {"column": "D", "column_keyword": "BRUT", "row_start": 10, "row_end": 30,
                          "section": "NOTE_9", "buckets": ["NOTE_9", "TITRES"]},
                          
    "note_11_cash": {"column": "D", "column_keyword": "BRUT", "row_start": 10, "row_end": 30,
                      "section": "NOTE_11", "buckets": ["NOTE_11", "CASH"]},

    # PASSIF NOTES
    "note_16_financial_debts": {"column": "E", "column_keyword": "NET", "row_start": 10, "row_end": 40,
                                 "section": "NOTE_16A", "buckets": ["NOTE_16", "DETTES_FIN"]},
                                 
    "note_16a_financial_debts": {"column": "E", "column_keyword": "NET", "row_start": 10, "row_end": 40,
                                 "section": "NOTE_16A", "buckets": ["NOTE_16A", "DETTES_FIN"]},
                                 
    "note_16b_suppliers": {"column": "E", "column_keyword": "NET", "row_start": 10, "row_end": 40,
                           "section": "NOTE_16B", "buckets": ["NOTE_16B", "FOURNISSEURS"]},
                           
    "note_17_suppliers": {"column": "E", "column_keyword": "NET", "row_start": 10, "row_end": 40,
                           "section": "NOTE_17", "buckets": ["NOTE_17", "FOURNISSEURS"]},
                           
    "note_18_fiscal_debts": {"column": "E", "column_keyword": "NET", "row_start": 10, "row_end": 40,
                           "section": "NOTE_18", "buckets": ["NOTE_18", "FISCAL"]},
                           
    "note_19_other_debts": {"column": "E", "column_keyword": "NET", "row_start": 10, "row_end": 30,
                      "section": "NOTE_19", "buckets": ["NOTE_19", "AUTRES_DETTES"]},
                      
    "note_20_bank_overdrafts": {"column": "E", "column_keyword": "NET", "row_start": 10, "row_end": 30,
                      "section": "NOTE_20", "buckets": ["NOTE_20", "BANQUES_PASSIF"]},
                      
    "note_27_turnover": {"column": "C", "column_keyword": "MONTANT", "row_start": 10, "row_end": 40,
                         "section": "NOTE_27A", "buckets": ["NOTE_27A", "CA"]},
                         
    "note_28_misc": {"column": "C", "column_keyword": "MONTANT", "row_start": 10, "row_end": 60,
                     "section": "NOTE_28", "buckets": ["NOTE_28", "DIVERS"]},

    # RSHEETS & NOTE 13 & PAGES
    "r1_ident": {"column": "B", "row_start": 5, "row_end": 60, "section": "R1", "buckets": ["R1"]},
    "r2_stat": {"column": "C", "row_start": 5, "row_end": 60, "section": "R2", "buckets": ["R2"]},
    "r3_stat": {"column": "C", "row_start": 5, "row_end": 60, "section": "R3", "buckets": ["R3"]},
    "note_13_capital": {"column": "C", "column_keyword": "Montant", "row_start": 5, "row_end": 40, 
                        "section": "NOTE_13", "buckets": ["NOTE_13", "CAPITAL"]},
    "page_garde": {"column": "C", "row_start": 5, "row_end": 60, "section": "PAGE_GARDE", "buckets": ["PAGE_GARDE"]},

    # NOTE: Class 9 and Class 8 usually map to Tableau de Flux or Notes
    # For now we map them to Flux D5-D10 or generic catch-all notes if possible
    # But usually they are handled by explicit manual rules or complex logic.
    # We will verify if we can dump them in Notes.
    # For this exercise, we map them to a generic bucket in Tableau Flux.
    "flux_ops": {"column": "E", "row_start": 10, "row_end": 20,
                 "section": "TABLEAU DES FLUX DE TRESORERIE", "buckets": ["FLUX", "OPERATIONS"]},
}

WINDOW_DEFAULT_SHEETS = {
    "bilan_asset_fixed": "BILAN",
    "bilan_asset_current": "BILAN",
    "bilan_equity": "BILAN",
    "bilan_liability": "BILAN",
    "cr_revenue": "COMPTE_RES",
    "cr_charge": "COMPTE_RES",
    "special_provisions": "BILAN",
    "special_other": "BILAN",
    "flux_ops": "TABLEAU DES FLUX DE TRESORERIE",
    "note_3a_fixed_assets": "NOTE 3A",
    "note_3c_amortization": "NOTE 3C",
    "note_4_stocks": "NOTE_4",
    "note_5_receivables": "NOTE_5",
    "note_6_securities": "NOTE_6",
    "note_16a_financial_debts": "NOTE_16A",
    "note_16b_suppliers": "NOTE_16B",
    "note_7_receivables": "NOTE_7",
    "note_6_stocks": "NOTE_6",
    "r1_ident": "R1",
    "r2_stat": "R2",
    "r3_stat": "R3",
    "note_13_capital": "NOTE13",
    "page_garde": "PAGE_GARDE",
}

# Mapping exact entre Libellés Excel (tout en bas) et Catégories
# Permet un ciblage ligne par ligne précis au lieu de fenêtres vagues.
EXCEL_LABEL_MAPPING = {
    # NOTE 3A / Bilan Actif Immob.
    "Frais de développement": "FA_INTANGIBLE",
    "Brevets, licences, logiciels": "FA_INTANGIBLE",
    "Fonds commercial": "FA_INTANGIBLE",
    
    # NOTE 4: Immo Financières
    "Titres de participation": "FA_FINANCIAL",
    "Prêts et créances": "FA_FINANCIAL_LOANS",
    "Titres immobilisés": "FA_FINANCIAL_OTHER",
    
    # NOTE 6: Stocks (Supposé)
    "Marchandises": "CA_INVENTORY",
    "Matières premières": "CA_INVENTORY",
    "Produits finis": "CA_INVENTORY", 
    "Stocks": "CA_INVENTORY",

    # NOTE 7: Créances (Supposé)
    "Clients": "CA_RECEIVABLES",
    "Clients douteux ou litigieux": "CA_RECEIVABLES",
    "Personnel": "CA_RECEIVABLES",
    "Etat": "CA_RECEIVABLES",
    "Débiteurs divers": "CA_RECEIVABLES",
    
    # NOTE 16: Emprunts
    "Emprunts": "LI_FINANCIAL_DEBTS",
    "Emprunts obligataires": "LI_FINANCIAL_DEBTS",
    "Emprunts et dettes auprès des établissements de crédit": "LI_FINANCIAL_DEBTS",

    # NOTE 27: Chiffre d'Affaires
    "Ventes de marchandises": "RE_SALES",
    "Travaux et services vendus": "RE_SALES",
    "Chiffre d'affaires": "RE_SALES",
    
    # NOTE 13: Capital
    "Montant total": "EQ_CAPITAL",
    "Apporteurs, capital non appelé": "EQ_UNCALLED",
    
    # Autres
    "Capital": "EQ_CAPITAL",
    "Réserves légales": "EQ_RESERVES",
}

CATEGORY_PREFIX_GROUPS = {
    # Special exact matches first (handled by detect logic)
    "FA_AMORTIZATION": ["bilan_asset_fixed", "note_3c_amortization"],
    "FA_DEPRECIATION": ["bilan_asset_fixed", "note_3c_amortization"],
    "FA_FINANCIAL": ["bilan_asset_fixed", "note_4_financial_assets"],
    
    # Fixed Assets
    "FA_": ["bilan_asset_fixed", "note_3a_fixed_assets"],
    
    # Current Assets
    "CA_INVENTORY": ["bilan_asset_current", "note_6_inventory"],
    "CA_RECEIVABLES": ["bilan_asset_current", "note_7_receivables"],
    "CA_OTHER": ["bilan_asset_current", "note_8_other_receivables"],
    "CA_PREPAID": ["bilan_asset_current", "note_8_other_receivables"], # Often grouped with other receivables
    "CA_SECURITIES": ["bilan_asset_current", "note_9_securities"],
    "CA_CASH": ["bilan_asset_current", "note_11_cash"],
    "CA_": ["bilan_asset_current"], # Default fallback for CA

    # Equity
    "EQ_CAPITAL": ["bilan_equity", "note_13_capital"],
    "EQ_": ["bilan_equity"],
    
    # Liabilities
    "LI_FINANCIAL_DEBTS": ["bilan_liability", "note_16_financial_debts"],
    "LI_DEBT_LT": ["bilan_liability", "note_16_financial_debts"],
    "LI_SUPPLIERS": ["bilan_liability", "note_17_suppliers"],
    "LI_TAXES": ["bilan_liability", "note_18_fiscal_debts"],
    "LI_OTHER": ["bilan_liability", "note_19_other_debts"],
    "LI_PERSONNEL": ["bilan_liability", "note_19_other_debts"], # Often grouped
    "LI_DEFERRED": ["bilan_liability", "note_19_other_debts"],
    "LI_BANK": ["bilan_liability", "note_20_bank_overdrafts"],
    "LI_": ["bilan_liability"],

    # Revenue & Expenses (CR + Notes if applicable)
    "RE_SALES": ["cr_revenue", "note_27_turnover"],
    "RE_": ["cr_revenue"],
    "EX_": ["cr_charge"],

    # Map newly supported prefixes / Special
    "SP_PROVISIONS": ["special_provisions"],
    "SP_": ["special_other"],
    "CT_": ["special_other"], 
    "AN_": ["flux_ops"], 
}


@dataclass
class RuleMaterial:
    """Intermediate representation used before building DSFRule objects."""

    category: str
    classe: Optional[str]
    prefixes: List[str]
    sheet_alias: str
    window_key: str


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate DSF rules from mappings")
    parser.add_argument(
        "--mapping",
        default="config/dsf_final_mapping.json",
        help="Chemin vers le JSON double mapping",
    )
    parser.add_argument(
        "--inventory",
        default="data/dsf_inventory.json",
        help="Chemin vers l'inventaire du template DSF",
    )
    parser.add_argument(
        "--output",
        default="config/dsf_rule_generated.yaml",
        help="Fichier de sortie (YAML ou JSON)",
    )
    parser.add_argument("--version", default="0.2.0", help="Version à inscrire dans la config")
    parser.add_argument("--template", default="templates/DSF Normal standard.xlsx")
    parser.add_argument("--verbose", action="store_true")
    return parser.parse_args()


def load_mapping(path: Path) -> Dict[str, object]:
    return json.loads(path.read_text(encoding="utf-8"))


def detect_window_keys(category: str) -> List[str]:
    """Returns a list of windows applicable to this category."""
    # sort by prefix length desc to match most specific first
    sorted_prefixes = sorted(CATEGORY_PREFIX_GROUPS.keys(), key=len, reverse=True)
    
    for prefix in sorted_prefixes:
        if category.startswith(prefix):
            windows = CATEGORY_PREFIX_GROUPS[prefix]
            if isinstance(windows, str):
                return [windows]
            return windows
    return []


# Overrides for standard OHADA prefixes if mapping is too restrictive
PREFIX_OVERRIDES = {
    "CA_RECEIVABLES": ["41"], # Clients
    "CA_INVENTORY": ["31", "32", "33", "34", "35", "36", "37", "38"],
    "FA_FINANCIAL": ["26", "27"],
    "LI_PERSONNEL": ["42", "43"],
    "LI_TAXES": ["44"],
    "LI_SUPPLIERS": ["40"],
    "LI_OTHER": ["45", "46", "47"],
    "CA_CASH": ["50", "51", "52", "53", "54", "55", "56", "57", "58"], # Class 5 is Treasury
    "LI_BANK": ["56"], # Discoveries
    "EQ_CAPITAL": ["10"],
    "EQ_RESERVES": ["11"],
    "EQ_RESULT": ["13"],
}

def build_rule_materials(mapping_data: Dict[str, object]) -> List[RuleMaterial]:
    materials: List[RuleMaterial] = []
    level2 = mapping_data.get("LEVEL_2_INTERMEDIATE_TO_DSF", {})
    info = mapping_data.get("INTERMEDIATE_CATEGORIES_INFO", {})

    for category, meta in info.items():
        # Use simple str(classe) logic
        classe = meta.get("classe")
        
        # Merge JSON prefixes with Manual Overrides
        json_prefixes = [p for p in meta.get("prefixes", []) if p.isdigit()]
        manual_prefixes = PREFIX_OVERRIDES.get(category, [])
        combined_prefixes = list(set(json_prefixes + manual_prefixes))
        
        # If we have manual overrides, we might cross class boundaries, so relax class constraint
        if manual_prefixes:
            classe = None 

        window_keys = detect_window_keys(category)
        
        if not window_keys:
            # logger.warning("Catégorie %s ignorée (pas de fenêtre mapping)", category)
            continue
            
        for window_key in window_keys:
            level2_entry = level2.get(category, {})
            sheet_alias = WINDOW_DEFAULT_SHEETS.get(window_key, level2_entry.get("sheet", "BILAN"))
            
            if sheet_alias not in SHEET_ALIASES:
                # logger.warning("Aucun alias de feuille pour %s (%s)", category, sheet_alias)
                continue
                
            materials.append(
                RuleMaterial(
                    category=category,
                    classe=str(classe) if classe else None,
                    prefixes=combined_prefixes,
                    sheet_alias=sheet_alias,
                    window_key=window_key,
                )
            )
    return materials


def validate_window(
    inventory: DSFInventory,
    sheet_name: str,
    column: str,
    row_start: int,
    row_end: int,
) -> Tuple[int, int]:
    """Ensure the requested window contains at least one writable field."""
    fields = list(
        inventory.iter_fields(
            sheet_name,
            column_letter=column,
            row_min=row_start,
            row_max=row_end,
        )
    )
    fields = [field for field in fields if not field.locked and not field.has_formula]
    if not fields:
        logger.warning(
            "Fenêtre vide: %s!col%s L%s-L%s → élargissement automatique",
            sheet_name,
            column,
            row_start,
            row_end,
        )
        fallback = [
            field
            for field in inventory.iter_fields(sheet_name, column_letter=column)
            if not field.locked and not field.has_formula
        ]
        if not fallback:
            # Downgrade to warning and fake a range to allow generation
            logger.warning("FORCE: Aucune cellule trouvée pour %s col %s. Utilisation aveugle de L%s-L%s", sheet_name, column, row_start, row_end)
            return row_start, row_end
            # raise ValueError(f"Aucune cellule disponible pour {sheet_name} colonne {column}")
        rows = [field.row for field in fallback]
        return min(rows), max(rows)
    rows = [field.row for field in fields]
    return min(rows), max(rows)


import unicodedata

def normalize_text(text: str) -> str:
    """Nettoie le texte: minuscule, sans accents, sans saut de ligne."""
    if not text:
        return ""
    text = text.lower().replace("\n", " ").strip()
    # Supprime les accents (ex: BÂTIMENT -> batiment)
    return ''.join(c for c in unicodedata.normalize('NFD', text) if unicodedata.category(c) != 'Mn')

# Cache pur l'indexation des feuilles (évite de rescanner 100 fois la même feuille)
_SHEET_INDEX_CACHE = {}

def get_sheet_index(inventory: DSFInventory, sheet_name: str) -> Dict[str, int]:
    """Construit un index inversé {texte_normalisé: row_number} pour une recherche O(1)."""
    if sheet_name in _SHEET_INDEX_CACHE:
        return _SHEET_INDEX_CACHE[sheet_name]
    
    index = {}
    # On scanne une seule fois tous les champs de la feuille
    for field in inventory.iter_fields(sheet_name):
        if field.label:
            norm_label = normalize_text(field.label)
            # On stocke le mapping (Attention: si doublon, on garde le premier ou dernier selon besoin)
            # Ici on garde le premier rencontré par sécurité ou on peut stocker une liste
            if norm_label not in index: 
                index[norm_label] = field.row
                
            # Optimisation: Indexer aussi des sous-parties de texte (pour mots clés)
            # Attention ça peut grossir l'index, mais utile pour "Matériel de transport"
            # Si le label est très long, on n'indexe que le début ? Non, restons simple pour l'instant.
    
    _SHEET_INDEX_CACHE[sheet_name] = index
    return index

def find_row_by_label(
    inventory: DSFInventory,
    sheet_name: str,
    category: str
) -> Optional[int]:
    """Cherche une ligne spécifique (Priorise correspondances EXACTES)."""
    
    # 1. Récupération des mots clés cibles (Mapping manuel)
    target_keywords = []
    for label_key, cat_val in EXCEL_LABEL_MAPPING.items():
        if cat_val == category:
            target_keywords.append(normalize_text(label_key))
    
    if not target_keywords:
        return None

    # 2. Utilisation de l'Index Inversé
    sheet_index = get_sheet_index(inventory, sheet_name)
    
    # Recherche exacte normalisée (Priorité absolue demandée par l'utilisateur)
    for kw in target_keywords:
        # Essai 1: Correspondance exacte sur le label complet
        if kw in sheet_index:
             return sheet_index[kw]
             
    # Recherche "Contient" (Fallback si l'exact échoue, mais on vérifie que c'est significatif)
    # L'utilisateur a demandé "titres complets", donc on est strict.
    # On autorise "contient" seulement si le mot clé est très long (> 10 chars) pour éviter les faux positifs.
    for kw in target_keywords:
        if len(kw) > 10:
             for label_text, row in sheet_index.items():
                if kw == label_text: # Redondant avec Essai 1 mais conservé par symétrie
                     return row
                # Si le label est "Frais de R&D (net)", et on cherche "Frais de R&D", ça passe
                # Mais si on cherche "Frais", ça ne passe pas à cause du len check
                if kw in label_text: 
                     return row

    return None


def find_column_letter_by_header(inventory: DSFInventory, sheet_name: str, keyword: str) -> Optional[str]:
    """Cherche la lettre de colonne par correspondance EXACTE (ou très proche) du header."""
    # Access raw payload exposed in updated DsfInventory
    if not hasattr(inventory, "raw_payload"):
        return None
        
    sheets = inventory.raw_payload.get("sheets", {})
    sheet_data = sheets.get(sheet_name)
    if not sheet_data:
        return None
    
    input_columns = sheet_data.get("input_columns", [])
    keyword_norm = normalize_text(keyword)
    
    best_col = None
    
    for col_info in input_columns:
        header = col_info.get("header_value", "")
        header_norm = normalize_text(header)
        
        # 1. Correspondance exacte
        if header_norm == keyword_norm:
            return col_info.get("letter")
            
        # 2. Correspondance "Contient" mais stricte (le mot clé doit être un token complet idéalement)
        if keyword_norm in header_norm:
            best_col = col_info.get("letter")
            
    return best_col # Fallback au contains si pas d'exact trouvé


def build_rules(
    materials: Iterable[RuleMaterial],
    inventory: DSFInventory,
    version: str,
    template_name: str,
) -> DSFRuleSet:
    rules: List[DSFRule] = []
    priority = 10

    for material in materials:
        window = CATEGORY_WINDOWS.get(material.window_key)
        if not window:
            logger.warning("Fenêtre non définie pour %s", material.category)
            continue
        sheet_name = SHEET_ALIASES[material.sheet_alias]
        
        # --- NEW: Try to find precise COLUMN by header keyword ---
        column = window["column"] # Default
        col_keyword = window.get("column_keyword")
        if col_keyword:
             found_col = find_column_letter_by_header(inventory, sheet_name, col_keyword)
             if found_col:
                 column = found_col
                 # logger.info(f"Colonne dynamique trouvée pour {sheet_name} ({col_keyword}) -> {column}")

        # --- NEW: Try to find precise row by label ---
        precise_row = find_row_by_label(inventory, sheet_name, material.category)
        if precise_row:
            row_start = precise_row
            row_end = precise_row
            notes_suffix = f" (Ciblage précis L{precise_row})"
        else:
            # Fallback to window
            row_start, row_end = validate_window(
                inventory,
                sheet_name,
                column,
                window["row_start"],
                window["row_end"],
            )
            notes_suffix = ""

        sign = SIGN_BY_CLASS.get(material.classe or "")
        filter_notes = f"Classe {material.classe} via {material.category}"
        target = DSFRuleTarget(
            sheet=sheet_name,
            column=column,
            exercice_column_type="exercice_n",
            row_start=row_start,
            row_end=row_end,
        )
        filters = AccountFilter(
            include_classes=[material.classe] if material.classe else [],
            include_prefixes=material.prefixes,
            require_sign=sign,
            notes=filter_notes,
        )
        label = f"{material.category.replace('_', ' ').title()}"
        rule = DSFRule(
            id=f"auto_{material.category.lower()}",
            label=label,
            section=window["section"],
            priority=priority,
            target=target,
            filters=filters,
            buckets=window["buckets"],
            notes=f"Mapping auto généré pour {material.category}",
        )
        rules.append(rule)
        priority += 10

    metadata = {
        "generated_rules": len(rules),
        "source_mapping": "dsf_final_mapping.json",
    }
    return DSFRuleSet(
        version=version,
        template_name=template_name,
        generated_from_inventory="data/dsf_inventory.json",
        rules=rules,
        metadata=metadata,
    )


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> None:
    args = parse_args()
    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO)

    mapping_path = Path(args.mapping)
    inventory_path = Path(args.inventory)
    output_path = Path(args.output)

    logger.info("Chargement du mapping %s", mapping_path)
    mapping_data = load_mapping(mapping_path)

    logger.info("Chargement de l'inventaire %s", inventory_path)
    inventory = DSFInventory.from_json(inventory_path)

    materials = build_rule_materials(mapping_data)
    logger.info("Catégories retenues: %s", len(materials))

    rule_set = build_rules(materials, inventory, args.version, args.template)
    save_rule_set(rule_set, output_path)
    logger.info("%s règles générées → %s", len(rule_set.rules), output_path)


if __name__ == "__main__":
    main()
