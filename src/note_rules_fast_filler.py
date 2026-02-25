#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Fast fillers for notes 19, 15A, 15B, 16A, C1-NOTE 17, C1-NOTE 25, C2-NOTE 25.
Uses SYSCOHADA prefixes and fixed rules (no fuzzy matching).
"""

from __future__ import annotations

import logging
import re
import unicodedata
from dataclasses import dataclass
from typing import Dict, List, Optional, Sequence, Tuple

logger = logging.getLogger(__name__)


@dataclass
class NoteRule:
    row_keywords: Sequence[str]
    account_prefixes: Sequence[str]


NOTE_RULES: Dict[str, List[NoteRule]] = {
    # NOTE 18: Dettes fiscales et sociales
    "NOTE 18": [
        NoteRule(["personnel", "avances"], ["42"]),
        NoteRule(["personnel", "remunerations"], ["42"]),
        NoteRule(["autres", "personnel"], ["42"]),
        NoteRule(["caisse", "securite"], ["43"]),
        NoteRule(["caisse", "retraite"], ["43"]),
        NoteRule(["organismes", "sociaux"], ["43"]),
        NoteRule(["etat", "impots", "benefices"], ["44"]),
        NoteRule(["etat", "impots", "taxes"], ["44"]),
        NoteRule(["etat", "tva"], ["44"]),
        NoteRule(["etat", "retenus"], ["44"]),
        NoteRule(["autres", "dettes", "etat"], ["44"]),
    ],
    # NOTE 19: Autres dettes et provisions pour risques a court terme
    "NOTE 19": [
        NoteRule(["organismes", "internationaux"], ["45", "46", "47", "48"]),
        NoteRule(["apporteurs"], ["45"]),
        NoteRule(["associes", "compte", "courant"], ["45"]),
        NoteRule(["associes", "dividendes"], ["45"]),
        NoteRule(["groupe", "comptes", "courants"], ["45"]),
        NoteRule(["autres", "dettes", "associes"], ["45", "46"]),
        NoteRule(["credits", "divers"], ["46"]),
        NoteRule(["obligataires"], ["16", "17"]),
        NoteRule(["remunerations", "administrateurs"], ["46"]),
        NoteRule(["factor"], ["46"]),
        NoteRule(["versements", "restants"], ["45", "46"]),
        NoteRule(["compte", "transitoire"], ["47"]),
        NoteRule(["autres", "crediteurs"], ["46"]),
    ],
    # NOTE 15A: Subventions et provisions reglementees
    "NOTE 15A": [
        NoteRule(["etat"], ["14"]),
        NoteRule(["regions"], ["14"]),
        NoteRule(["departements"], ["14"]),
        NoteRule(["communes"], ["14"]),
        NoteRule(["entites", "publiques"], ["14"]),
        NoteRule(["entites", "organismes", "prives"], ["14"]),
        NoteRule(["organismes", "internationaux"], ["14"]),
        NoteRule(["autres"], ["14"]),
        NoteRule(["amortissements", "derogatoires"], ["15"]),
        NoteRule(["plus", "value", "cession"], ["15"]),
        NoteRule(["provisions", "speciales"], ["15"]),
        NoteRule(["provisions", "immobilisations"], ["15"]),
        NoteRule(["provisions", "stocks"], ["15"]),
    ],
    # NOTE 15B: Autres fonds propres
    "NOTE 15B": [
        NoteRule(["titres", "participatifs"], ["19"]),
        NoteRule(["avances", "conditionnees"], ["19"]),
        NoteRule(["titres", "subordonnes"], ["19"]),
        NoteRule(["obligations", "remboursables"], ["19"]),
        NoteRule(["autres"], ["19"]),
    ],
    # NOTE 16A: Dettes financieres et ressources assimilees
    "NOTE 16A": [
        NoteRule(["emprunts", "obligatoires"], ["16"]),
        NoteRule(["emprunts", "dettes", "etablissements"], ["16"]),
        NoteRule(["avances", "etat"], ["16"]),
        NoteRule(["avances", "comptes", "bloques"], ["16"]),
        NoteRule(["depots", "cautionnement"], ["16"]),
        NoteRule(["interets", "courus"], ["16"]),
        NoteRule(["avances", "associes"], ["16"]),
        NoteRule(["autres", "emprunts"], ["16"]),
        NoteRule(["dettes", "participations"], ["16"]),
        NoteRule(["comptes", "permanents", "bloques"], ["16"]),
        NoteRule(["credit", "bail", "immobilier"], ["17"]),
        NoteRule(["credit", "bail", "mobilier"], ["17"]),
        NoteRule(["location", "vente"], ["17"]),
    ],
    # NOTE 27A: Charges de personnel
    "NOTE 27A": [
        NoteRule(["remunerations", "directes"], ["64"]),
        NoteRule(["indemnites"], ["64"]),
        NoteRule(["charges", "sociales"], ["64", "43"]),
        NoteRule(["exploitant"], ["64"]),
        NoteRule(["personnel", "exterieur"], ["64"]),
        NoteRule(["autres", "charges", "sociales"], ["64", "43"]),
    ],
    # NOTE 28: Provisions et depreciations inscrites au bilan
    "NOTE 28": [
        NoteRule(["provisions", "reglementes"], ["15"]),
        NoteRule(["provisions", "financieres"], ["15"]),
        NoteRule(["depreciation", "immobilisations"], ["29"]),
        NoteRule(["depreciations", "stocks"], ["39"]),
        NoteRule(["depreciations", "fournisseurs"], ["49"]),
        NoteRule(["depreciations", "clients"], ["49"]),
        NoteRule(["depreciations", "creances"], ["49"]),
        NoteRule(["depreciations", "titres"], ["59"]),
        NoteRule(["depreciations", "valeurs"], ["59"]),
        NoteRule(["depreciations", "disponibilite"], ["59"]),
        NoteRule(["provisions", "court", "termes"], ["15"]),
    ],
    # NOTE 30: Autres charges et produits HAO
    "NOTE 30": [
        NoteRule(["charges", "hao"], ["67"]),
        NoteRule(["pertes", "creances", "hao"], ["67"]),
        NoteRule(["dons", "liberalites", "accords"], ["67"]),
        NoteRule(["abandons", "creances"], ["67"]),
        NoteRule(["charges", "provisionnees"], ["67"]),
        NoteRule(["dotations", "hors"], ["67"]),
        NoteRule(["participation", "travailleurs"], ["67"]),
        NoteRule(["subventions", "equilibre"], ["67"]),
        NoteRule(["produits", "hao"], ["77"]),
        NoteRule(["dons", "liberalites", "obtenus"], ["77"]),
    ],
    # NOTE 32: Production de l'exercice (produits)
    "NOTE 32": [],
    # C1-NOTE 17: Extrait balance fournisseurs (compte en colonne A)
    "C1-NOTE 17": [],
    # C1-NOTE 25: Synthese impots et taxes verses
    "C1-NOTE 25": [
        NoteRule(["impots", "societ"], ["69"]),
        NoteRule(["irpp"], ["44", "63"]),
        NoteRule(["traitements", "salaires"], ["64"]),
        NoteRule(["ircm"], ["44", "63"]),
        NoteRule(["revenus", "fonciers"], ["63"]),
        NoteRule(["benefice", "artisanaux"], ["63"]),
        NoteRule(["benefice", "agricole"], ["63"]),
        NoteRule(["professions", "non", "commerciales"], ["63"]),
        NoteRule(["revenus", "non", "commerciaux"], ["63"]),
        NoteRule(["taxe", "valeur", "ajoute"], ["445", "44"]),
        NoteRule(["droits", "accises"], ["44", "63"]),
    ],
    # C2-NOTE 25: Tableau regularisation droits d'accises
    "C2-NOTE 25": [],
    # C1-NOTE 28: Tableau recapitulatif fiscal des provisions
    "C1-NOTE 28": [
        NoteRule(["provisions", "reglementes"], ["15"]),
        NoteRule(["provisions", "financieres"], ["15"]),
        NoteRule(["depreciation", "immobilisations"], ["29"]),
        NoteRule(["depreciations", "stocks"], ["39"]),
        NoteRule(["depreciations", "fournisseurs"], ["49"]),
        NoteRule(["depreciations", "clients"], ["49"]),
        NoteRule(["depreciations", "creances"], ["49"]),
        NoteRule(["depreciations", "titres"], ["59"]),
    ],
    # C2-NOTE 28: Tableau recapitulatif fiscal des provisions
    "C2-NOTE 28": [
        NoteRule(["provisions", "reglementes"], ["15"]),
        NoteRule(["provisions", "financieres"], ["15"]),
        NoteRule(["depreciation", "immobilisations"], ["29"]),
        NoteRule(["depreciations", "stocks"], ["39"]),
        NoteRule(["depreciations", "fournisseurs"], ["49"]),
        NoteRule(["depreciations", "clients"], ["49"]),
        NoteRule(["depreciations", "creances"], ["49"]),
        NoteRule(["depreciations", "titres"], ["59"]),
        NoteRule(["depreciations", "disponibilite"], ["59"]),
        NoteRule(["provisions", "court", "termes"], ["15"]),
    ],
}


def _normalize(text: str) -> str:
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.lower()
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _detect_header_row(ws) -> Optional[int]:
    for row_num in range(1, min(30, ws.max_row + 1)):
        value = ws.cell(row_num, 1).value
        if not value:
            continue
        if "libell" in _normalize(str(value)):
            return row_num
    return None


def _detect_value_columns(ws, header_row: Optional[int]) -> Tuple[Optional[int], Optional[int]]:
    if not header_row:
        return None, None

    header_map = {}
    for col_num in range(2, min(15, ws.max_column + 1)):
        cell_val = ws.cell(header_row, col_num).value
        if cell_val:
            header_map[col_num] = _normalize(str(cell_val))

    col_n1 = None
    col_n = None
    for col_num, label in header_map.items():
        if "n-1" in label or "n1" in label:
            col_n1 = col_num
        elif "n" in label or "exercice" in label:
            col_n = col_num

    if col_n1 is None or col_n is None:
        data_cols = sorted(header_map.keys())
        if len(data_cols) >= 2:
            col_n1 = col_n1 or data_cols[0]
            col_n = col_n or data_cols[-1]
        elif len(data_cols) == 1:
            col_n = col_n or data_cols[0]

    return col_n, col_n1


def _is_formula(cell) -> bool:
    return isinstance(cell.value, str) and cell.value.startswith("=") or cell.data_type == "f"


def _sum_accounts(
    accounts: Dict[str, object],
    accounts_n1: Optional[Dict[str, object]],
    prefixes: Sequence[str],
    keywords: Sequence[str],
) -> Tuple[float, float]:
    prefixes = [p for p in prefixes if p]
    keywords_norm = [_normalize(k) for k in keywords if k]

    total_n = 0.0
    total_n1 = 0.0

    for compte, row in accounts.items():
        compte_str = str(compte)
        if prefixes and not any(compte_str.startswith(p) for p in prefixes):
            continue
        label_norm = _normalize(getattr(row, "label", ""))
        if keywords_norm and not all(k in label_norm for k in keywords_norm):
            continue
        total_n += float(getattr(row, "solde_final", 0) or 0)

    if accounts_n1:
        for compte, row in accounts_n1.items():
            compte_str = str(compte)
            if prefixes and not any(compte_str.startswith(p) for p in prefixes):
                continue
            label_norm = _normalize(getattr(row, "label", ""))
            if keywords_norm and not all(k in label_norm for k in keywords_norm):
                continue
            total_n1 += float(getattr(row, "solde_final", 0) or 0)

    return total_n, total_n1


class NoteRulesFastFiller:
    def __init__(self, workbook, accounts: Dict[str, object], accounts_n1: Optional[Dict[str, object]] = None):
        self.wb = workbook
        self.accounts = accounts
        self.accounts_n1 = accounts_n1 or {}

        self._label_index = {
            compte: _normalize(getattr(row, "label", ""))
            for compte, row in accounts.items()
        }

    def fill_note(self, sheet_name: str) -> int:
        if sheet_name not in self.wb.sheetnames:
            return 0

        ws = self.wb[sheet_name]
        header_row = _detect_header_row(ws)
        col_n, col_n1 = _detect_value_columns(ws, header_row)

        if col_n is None:
            col_n = 2
        if col_n1 is None:
            col_n1 = 3

        rules = NOTE_RULES.get(sheet_name, [])
        if not rules and sheet_name not in {"C1-NOTE 17", "NOTE 32"}:
            logger.info("No rules configured for %s", sheet_name)
            return 0

        assignments = 0

        for row_num in range(1, ws.max_row + 1):
            label_cell = ws.cell(row_num, 1).value
            if not label_cell:
                continue

            label_norm = _normalize(str(label_cell))

            # NOTE 32: match production rows by account label contains row label
            if sheet_name == "NOTE 32":
                if label_norm in {"designation de produits", "libelles", "libelle", "total"}:
                    continue
                if "total" in label_norm:
                    continue
                if len(label_norm) < 4:
                    continue

                value_n, value_n1 = _sum_accounts(
                    self.accounts,
                    self.accounts_n1,
                    ["70"],
                    [label_norm],
                )

                if col_n and not _is_formula(ws.cell(row_num, col_n)):
                    ws.cell(row_num, col_n).value = value_n
                    assignments += 1
                if col_n1 and not _is_formula(ws.cell(row_num, col_n1)):
                    ws.cell(row_num, col_n1).value = value_n1
                    assignments += 1
                continue

            # C1-NOTE 17: numeric account in column A
            if sheet_name == "C1-NOTE 17":
                if label_norm.isdigit():
                    compte = label_norm
                    row_n = self.accounts.get(compte)
                    row_n1 = self.accounts_n1.get(compte)
                    if row_n:
                        if col_n and not _is_formula(ws.cell(row_num, col_n)):
                            ws.cell(row_num, col_n).value = float(row_n.solde_final or 0)
                            assignments += 1
                        if col_n1 and row_n1 and not _is_formula(ws.cell(row_num, col_n1)):
                            ws.cell(row_num, col_n1).value = float(row_n1.solde_final or 0)
                            assignments += 1
                continue

            for rule in rules:
                if all(_normalize(k) in label_norm for k in rule.row_keywords):
                    value_n, value_n1 = _sum_accounts(
                        self.accounts,
                        self.accounts_n1,
                        rule.account_prefixes,
                        rule.row_keywords,
                    )

                    if col_n and not _is_formula(ws.cell(row_num, col_n)):
                        ws.cell(row_num, col_n).value = value_n
                        assignments += 1
                    if col_n1 and not _is_formula(ws.cell(row_num, col_n1)):
                        ws.cell(row_num, col_n1).value = value_n1
                        assignments += 1
                    break

        logger.info("Fast rules filled %s: %s assignments", sheet_name, assignments)
        return assignments

    def fill_all(self, sheets: Sequence[str]) -> Dict[str, int]:
        results: Dict[str, int] = {}
        for sheet in sheets:
            results[sheet] = self.fill_note(sheet)
        return results
