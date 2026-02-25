# -*- coding: utf-8 -*-
"""
Module Tableau des Flux de Trésorerie SYSCOHADA — Priorité 2 (A4)
Calcule et remplit le TFT par la méthode indirecte (standard SYSCOHADA).
"""
from __future__ import annotations

import logging
from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import Path
from typing import Dict, List, Optional

logger = logging.getLogger(__name__)


def _dec(v) -> Decimal:
    if v is None:
        return Decimal("0")
    try:
        return Decimal(str(v))
    except Exception:
        return Decimal("0")


# ---------------------------------------------------------------------------
# Mapping comptes SYSCOHADA → catégories de flux
# ---------------------------------------------------------------------------

# Comptes d'amortissements et provisions (à réintégrer dans flux opérationnels)
AMORT_PREFIXES = ["28", "29", "39", "49", "59"]

# Variation BFR — Actif circulant (augmentation = emploi = signe -)
BFR_ACTIF_PREFIXES = ["31", "32", "33", "34", "35", "36", "37", "38",
                       "411", "412", "413", "416", "417", "461", "462",
                       "445", "4451", "4452", "476"]

# Variation BFR — Passif circulant (augmentation = ressource = signe +)
BFR_PASSIF_PREFIXES = ["401", "402", "404", "408", "419",
                        "421", "422", "431", "441", "444",
                        "471", "472", "477", "478"]

# Investissements — Acquisitions (emploi = signe -)
INVEST_ACQUI_PREFIXES = ["20", "21", "22", "23", "24", "25", "26", "27"]

# Investissements — Cessions (ressource = signe +)
INVEST_CESSION_PREFIXES = ["462"]  # Créances sur cessions d'immobilisations

# Financement — Ressources
FINANCE_RESSOURCES_PREFIXES = ["101", "102", "161", "162", "163", "164"]

# Financement — Emplois (remboursements, dividendes)
FINANCE_EMPLOIS_PREFIXES = ["121", "129"]  # Report à nouveau (dividendes distribués)

# Trésorerie
TRESORERIE_PREFIXES = ["511", "512", "521", "522", "523", "524", "571", "572", "573", "581",
                        "500", "501", "502", "503", "504"]
TRESORERIE_PASSIF_PREFIXES = ["561", "565"]  # Concours bancaires (passif)

# Résultat net
RESULTAT_PREFIXES = ["131", "132"]


# ---------------------------------------------------------------------------
# Dataclasses
# ---------------------------------------------------------------------------

@dataclass
class CashFlowSection:
    """Section du Tableau des Flux de Trésorerie."""
    name: str
    items: Dict[str, Decimal] = field(default_factory=dict)

    @property
    def total(self) -> Decimal:
        return sum(self.items.values(), Decimal("0"))


@dataclass
class CashFlowStatement:
    """Tableau des Flux de Trésorerie complet."""
    operating: CashFlowSection      # FA — Flux opérationnels
    investing: CashFlowSection      # FB — Flux d'investissement
    financing: CashFlowSection      # FC — Flux de financement
    tresorerie_ouverture: Decimal = Decimal("0")
    tresorerie_cloture: Decimal = Decimal("0")

    @property
    def variation_nette(self) -> Decimal:
        return self.operating.total + self.investing.total + self.financing.total

    @property
    def is_balanced(self) -> bool:
        """Vérifie que FA+FB+FC = ΔTrésorerie (tolérance 1000 FCFA)."""
        delta_tresorerie = self.tresorerie_cloture - self.tresorerie_ouverture
        return abs(self.variation_nette - delta_tresorerie) <= Decimal("1000")


# ---------------------------------------------------------------------------
# Calculateur
# ---------------------------------------------------------------------------

class DSFCashFlowFiller:
    """
    Calcule et remplit le Tableau des Flux de Trésorerie SYSCOHADA.
    Utilise la méthode indirecte (résultat net + retraitements).

    Structure SYSCOHADA :
        FA = Résultat net + Amortissements/Provisions - ΔBFR
        FB = Acquisitions - Cessions d'immobilisations
        FC = Augmentations capital/emprunts - Remboursements/Dividendes
        ΔTrésorerie = FA + FB + FC
    """

    def __init__(self, wb, normalized_rows: list, previous_rows: Optional[list] = None):
        """
        Args:
            wb: Workbook openpyxl du DSF
            normalized_rows: Balance N normalisée (NormalizedBalanceRow)
            previous_rows: Balance N-1 normalisée (optionnel)
        """
        self.wb = wb
        self._idx_n  = self._build_index(normalized_rows or [])
        self._idx_n1 = self._build_index(previous_rows or [])

    # ------------------------------------------------------------------
    # API publique
    # ------------------------------------------------------------------

    def compute(self) -> CashFlowStatement:
        """Calcule le TFT complet."""
        operating = self._compute_operating()
        investing  = self._compute_investing()
        financing  = self._compute_financing()
        tresorerie_ouverture = self._sum_tresorerie(self._idx_n1)
        tresorerie_cloture   = self._sum_tresorerie(self._idx_n)

        stmt = CashFlowStatement(
            operating=operating,
            investing=investing,
            financing=financing,
            tresorerie_ouverture=tresorerie_ouverture,
            tresorerie_cloture=tresorerie_cloture,
        )

        logger.info(
            "TFT calculé : FA=%s / FB=%s / FC=%s / ΔTréso=%s / Équilibré=%s",
            operating.total, investing.total, financing.total,
            stmt.variation_nette, stmt.is_balanced,
        )
        return stmt

    def fill_worksheet(self) -> int:
        """
        Écrit le TFT dans la feuille correspondante du workbook.

        Returns:
            Nombre de cellules écrites
        """
        ws = self._find_tft_sheet()
        if ws is None:
            logger.warning("Feuille Tableau des Flux de Trésorerie introuvable — TFT non écrit")
            return 0

        stmt = self.compute()
        written = 0

        # Mapping libellés → valeurs à écrire
        values_to_write = {
            # FA — Opérationnel
            "Résultat net de l'exercice":           stmt.operating.items.get("resultat_net", Decimal("0")),
            "Dotations aux amortissements":          stmt.operating.items.get("amortissements", Decimal("0")),
            "Variation du besoin en fonds de roulement": stmt.operating.items.get("delta_bfr", Decimal("0")),
            "Flux net des activités opérationnelles": stmt.operating.total,
            # FB — Investissement
            "Acquisitions d'immobilisations":        stmt.investing.items.get("acquisitions", Decimal("0")),
            "Cessions d'immobilisations":            stmt.investing.items.get("cessions", Decimal("0")),
            "Flux net des activités d'investissement": stmt.investing.total,
            # FC — Financement
            "Augmentations de capital":              stmt.financing.items.get("augmentations_capital", Decimal("0")),
            "Nouveaux emprunts":                     stmt.financing.items.get("nouveaux_emprunts", Decimal("0")),
            "Remboursements et dividendes":          stmt.financing.items.get("remboursements", Decimal("0")),
            "Flux net des activités de financement": stmt.financing.total,
            # Synthèse
            "Variation nette de trésorerie":         stmt.variation_nette,
            "Trésorerie à l'ouverture":              stmt.tresorerie_ouverture,
            "Trésorerie à la clôture":               stmt.tresorerie_cloture,
        }

        for label, value in values_to_write.items():
            row_idx = self._find_row_by_label(ws, label)
            if row_idx:
                try:
                    # Écrire dans la colonne de données (typiquement C ou D)
                    ws[f"C{row_idx}"] = float(value)
                    written += 1
                except Exception as exc:
                    logger.debug("Erreur écriture TFT '%s' : %s", label, exc)

        logger.info("TFT : %d cellules écrites dans '%s'", written, ws.title)
        return written

    # ------------------------------------------------------------------
    # Calculs par section
    # ------------------------------------------------------------------

    def _compute_operating(self) -> CashFlowSection:
        """FA — Flux des activités opérationnelles (méthode indirecte)."""
        # 1. Résultat net
        resultat_net = self._sum_prefixes(self._idx_n, RESULTAT_PREFIXES)

        # 2. Réintégration amortissements et provisions (charges non décaissées)
        amortissements = abs(self._sum_prefixes(self._idx_n, AMORT_PREFIXES))

        # 3. Variation BFR = (BFR_N - BFR_N1)
        bfr_actif_n  = self._sum_prefixes(self._idx_n,  BFR_ACTIF_PREFIXES)
        bfr_actif_n1 = self._sum_prefixes(self._idx_n1, BFR_ACTIF_PREFIXES)
        bfr_passif_n  = self._sum_prefixes(self._idx_n,  BFR_PASSIF_PREFIXES)
        bfr_passif_n1 = self._sum_prefixes(self._idx_n1, BFR_PASSIF_PREFIXES)

        bfr_n  = bfr_actif_n  - bfr_passif_n
        bfr_n1 = bfr_actif_n1 - bfr_passif_n1
        delta_bfr = -(bfr_n - bfr_n1)  # Augmentation BFR = emploi = négatif

        return CashFlowSection(
            name="Activités opérationnelles",
            items={
                "resultat_net":    resultat_net,
                "amortissements":  amortissements,
                "delta_bfr":       delta_bfr,
            },
        )

    def _compute_investing(self) -> CashFlowSection:
        """FB — Flux des activités d'investissement."""
        # Acquisitions : variation des immobilisations brutes (augmentation = emploi)
        immo_n  = self._sum_prefixes(self._idx_n,  INVEST_ACQUI_PREFIXES)
        immo_n1 = self._sum_prefixes(self._idx_n1, INVEST_ACQUI_PREFIXES)
        acquisitions = -(immo_n - immo_n1)  # Augmentation = emploi = négatif

        # Cessions : créances sur cessions d'immobilisations
        cessions = self._sum_prefixes(self._idx_n, INVEST_CESSION_PREFIXES)

        return CashFlowSection(
            name="Activités d'investissement",
            items={
                "acquisitions": acquisitions,
                "cessions":     cessions,
            },
        )

    def _compute_financing(self) -> CashFlowSection:
        """FC — Flux des activités de financement."""
        # Ressources : augmentations capital + nouveaux emprunts
        capital_n  = self._sum_prefixes(self._idx_n,  ["101", "102"])
        capital_n1 = self._sum_prefixes(self._idx_n1, ["101", "102"])
        augmentations_capital = capital_n - capital_n1

        emprunts_n  = self._sum_prefixes(self._idx_n,  ["161", "162", "163", "164"])
        emprunts_n1 = self._sum_prefixes(self._idx_n1, ["161", "162", "163", "164"])
        nouveaux_emprunts = emprunts_n - emprunts_n1

        # Emplois : remboursements et dividendes distribués (variation report à nouveau)
        rna_n  = self._sum_prefixes(self._idx_n,  FINANCE_EMPLOIS_PREFIXES)
        rna_n1 = self._sum_prefixes(self._idx_n1, FINANCE_EMPLOIS_PREFIXES)
        remboursements = -(rna_n - rna_n1)  # Diminution RNA = dividendes distribués

        return CashFlowSection(
            name="Activités de financement",
            items={
                "augmentations_capital": augmentations_capital,
                "nouveaux_emprunts":     nouveaux_emprunts,
                "remboursements":        remboursements,
            },
        )

    def _sum_tresorerie(self, index: Dict[str, Decimal]) -> Decimal:
        """Calcule la trésorerie nette (actif - passif bancaire)."""
        actif  = self._sum_prefixes(index, TRESORERIE_PREFIXES)
        passif = self._sum_prefixes(index, TRESORERIE_PASSIF_PREFIXES)
        return actif - passif

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _build_index(rows: list) -> Dict[str, Decimal]:
        """Construit un index {compte: solde_net} depuis les NormalizedBalanceRow."""
        index: Dict[str, Decimal] = {}
        for row in rows:
            try:
                compte = str(getattr(row, "compte", "") or "").strip()
                closing = getattr(row, "closing_balance", None)
                if closing is not None:
                    solde = _dec(closing)
                else:
                    debit  = _dec(getattr(row, "debit_total",  0))
                    credit = _dec(getattr(row, "credit_total", 0))
                    solde  = debit - credit
                if compte:
                    index[compte] = index.get(compte, Decimal("0")) + solde
            except Exception as exc:
                logger.debug("Erreur indexation TFT compte %s : %s", getattr(row, "compte", "?"), exc)
        return index

    @staticmethod
    def _sum_prefixes(index: Dict[str, Decimal], prefixes: List[str]) -> Decimal:
        """Somme les comptes dont le numéro commence par l'un des préfixes."""
        total = Decimal("0")
        for compte, solde in index.items():
            for p in prefixes:
                if compte.startswith(p):
                    total += solde
                    break
        return total

    def _find_tft_sheet(self):
        """Cherche la feuille TFT dans le workbook."""
        keywords = ["FLUX", "TRESORERIE", "TFT", "CASH FLOW"]
        for name in self.wb.sheetnames:
            name_up = name.upper()
            if any(kw in name_up for kw in keywords):
                return self.wb[name]
        return None

    @staticmethod
    def _find_row_by_label(ws, label: str) -> Optional[int]:
        """Cherche la ligne contenant le libellé (insensible à la casse)."""
        label_lower = label.lower()
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
            for cell in row:
                if isinstance(cell.value, str) and label_lower in cell.value.lower():
                    return cell.row
        return None


# ---------------------------------------------------------------------------
# Fonction utilitaire pour le pipeline
# ---------------------------------------------------------------------------

def fill_cash_flow(
    wb,
    normalized_rows: list,
    previous_rows: Optional[list] = None,
) -> int:
    """
    Calcule et remplit le Tableau des Flux de Trésorerie dans le workbook DSF.

    Args:
        wb: Workbook openpyxl (déjà chargé)
        normalized_rows: Balance N normalisée
        previous_rows: Balance N-1 normalisée (optionnel)

    Returns:
        Nombre de cellules écrites
    """
    filler = DSFCashFlowFiller(wb, normalized_rows, previous_rows)
    return filler.fill_worksheet()


__all__ = [
    "DSFCashFlowFiller",
    "CashFlowStatement",
    "CashFlowSection",
    "fill_cash_flow",
]
