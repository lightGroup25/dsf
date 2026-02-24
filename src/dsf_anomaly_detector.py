# -*- coding: utf-8 -*-
"""
P3-MF4 : Détection des soldes anormaux dans la balance SYSCOHADA.
Signale les comptes dont le sens du solde (Débit/Crédit) est contraire à la norme SYSCOHADA.
Seuil : solde anormal > 1 000 FCFA pour éviter le bruit sur les petits écarts.
"""
from __future__ import annotations

import logging
from dataclasses import dataclass, field
from decimal import Decimal
from typing import Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Table des sens normaux par tranche de numéro de compte SYSCOHADA
# (D = normalement Débiteur, C = normalement Créditeur)
# ---------------------------------------------------------------------------
EXPECTED_NORMAL_SIDE: List[Tuple[str, str, str]] = [
    # (préfixe_min, préfixe_max, côté_normal)
    ("10", "19", "C"),   # Classe 1 — Capitaux propres et dettes financières
    ("20", "29", "D"),   # Classe 2 — Immobilisations (Brut)
    ("28", "29", "C"),   # Classe 2 — Amortissements et provisions (exception)
    ("30", "39", "D"),   # Classe 3 — Stocks
    ("40", "40", "C"),   # 40x — Fournisseurs (créditeurs normaux)
    ("409", "409", "D"), # 409 — Fournisseurs débiteurs (exception)
    ("41", "41", "D"),   # 41x — Clients (débiteurs normaux)
    ("419", "419", "C"), # 419 — Clients créditeurs (exception)
    ("42", "48", "C"),   # 42x-48x — Dettes sociales/fiscales
    ("49", "49", "C"),   # 49x — Provisions sur créances
    ("50", "59", "D"),   # Classe 5 — Trésorerie (sauf 56x CBC)
    ("56", "56", "C"),   # 56x — Concours bancaires courants (créditeurs)
    ("60", "69", "D"),   # Classe 6 — Charges
    ("70", "79", "C"),   # Classe 7 — Produits
    ("80", "89", "D"),   # Classe 8 — Résultat HAO
]

ANOMALY_THRESHOLD = Decimal("1000")   # Seuil 1 000 FCFA


@dataclass
class AnomalyReport:
    """Rapport d'anomalie de solde."""
    compte: str
    label: str
    normal_side: str          # "D" ou "C"
    actual_side: str          # "D" ou "C"
    abnormal_amount: Decimal  # montant anormal (toujours positif)
    severity: str             # "INFO", "WARNING", "ERROR"

    def __str__(self) -> str:
        return (
            f"[{self.severity}] Compte {self.compte} ({self.label}) : "
            f"solde {self.actual_side} de {self.abnormal_amount:,.0f} FCFA "
            f"alors que le côté normal est {self.normal_side}"
        )


class DSFAnomalyDetector:
    """
    P3-MF4 : Détecte les soldes anormaux dans la balance normalisée.

    Compare le solde de chaque compte à son côté normal SYSCOHADA.
    Si le solde est du mauvais côté et > 1 000 FCFA, génère un rapport.
    """

    def __init__(self, tolerance: Decimal = ANOMALY_THRESHOLD):
        self.tolerance = tolerance

    def detect_abnormal_balances(self, balance_rows: list) -> List[AnomalyReport]:
        """
        Analyse tous les comptes de la balance et retourne la liste des anomalies.

        Args:
            balance_rows: Liste de NormalizedBalanceRow (balance N)

        Returns:
            Liste d'AnomalyReport, triée par sévérité puis par numéro de compte
        """
        reports: List[AnomalyReport] = []
        for row in balance_rows:
            try:
                report = self._check_row(row)
                if report is not None:
                    reports.append(report)
            except Exception as exc:
                logger.debug("Erreur analyse compte %s : %s", getattr(row, "compte", "?"), exc)

        # Trier : ERROR > WARNING > INFO, puis par compte
        severity_order = {"ERROR": 0, "WARNING": 1, "INFO": 2}
        reports.sort(key=lambda r: (severity_order.get(r.severity, 9), r.compte))

        if reports:
            logger.warning(
                "P3-MF4 : %d solde(s) anormal/aux détecté(s) (%d ERROR, %d WARNING)",
                len(reports),
                sum(1 for r in reports if r.severity == "ERROR"),
                sum(1 for r in reports if r.severity == "WARNING"),
            )
        else:
            logger.info("P3-MF4 : Aucun solde anormal détecté.")

        return reports

    def _check_row(self, row) -> Optional[AnomalyReport]:
        """Vérifie un compte individuel. Retourne un rapport si anomalie détectée."""
        compte = str(getattr(row, "compte", "") or "").strip()
        if not compte:
            return None

        # Ignorer les comptes de classe 9 (analytique, hors scope)
        if compte.startswith("9"):
            return None

        expected_side = self._get_expected_side(compte)
        if expected_side is None:
            return None  # Pas de règle définie pour ce compte

        # Calcul du solde net
        closing = getattr(row, "closing_balance", None)
        if closing is not None:
            solde = Decimal(str(closing))
        else:
            debit  = Decimal(str(getattr(row, "debit_total",  0) or 0))
            credit = Decimal(str(getattr(row, "credit_total", 0) or 0))
            solde  = debit - credit

        # Détecter le côté réel
        actual_side = "D" if solde > 0 else "C" if solde < 0 else None
        if actual_side is None:
            return None  # Solde nul = aucune anomalie

        abnormal_amount = abs(solde)
        if actual_side == expected_side:
            return None  # Côté correct

        # Solde du mauvais côté — vérifier le seuil
        if abnormal_amount < self.tolerance:
            return None  # En dessous du seuil de signalement

        label = str(getattr(row, "label", "") or "").strip()

        # Déterminer la sévérité selon la classe de compte
        severity = self._severity_for_account(compte, abnormal_amount)

        return AnomalyReport(
            compte=compte,
            label=label,
            normal_side=expected_side,
            actual_side=actual_side,
            abnormal_amount=abnormal_amount,
            severity=severity,
        )

    @staticmethod
    def _get_expected_side(compte: str) -> Optional[str]:
        """Retourne le côté normal (D/C) pour un compte donné."""
        # Chercher la règle la plus spécifique (préfixe le plus long)
        best_match: Optional[Tuple[str, str]] = None
        best_len = 0
        for prefix_min, prefix_max, side in EXPECTED_NORMAL_SIDE:
            len_min = len(prefix_min)
            prefix_compte = compte[:len_min]
            if prefix_min <= prefix_compte <= prefix_max:
                if len_min > best_len:
                    best_match = (prefix_min, side)
                    best_len = len_min
        return best_match[1] if best_match else None

    @staticmethod
    def _severity_for_account(compte: str, amount: Decimal) -> str:
        """Détermine la sévérité selon la classe de compte et le montant."""
        # Classe 1 et 7 : anomalies de capitaux/produits → ERROR si > 100k
        if compte[:1] in ("1", "7") and amount > Decimal("100000"):
            return "ERROR"
        # Classe 6 : charges du mauvais côté → WARNING
        if compte[:1] == "6":
            return "WARNING"
        # Classe 4 : dettes du mauvais côté → WARNING
        if compte[:1] == "4" and amount > Decimal("10000"):
            return "WARNING"
        return "INFO"

    def log_report(self, reports: List[AnomalyReport]) -> None:
        """Journalise le rapport d'anomalies de manière lisible."""
        if not reports:
            logger.info("P3-MF4 : Aucune anomalie de solde.")
            return
        logger.warning("=== P3-MF4 : Rapport des soldes anormaux (%d) ===", len(reports))
        for r in reports:
            if r.severity == "ERROR":
                logger.error("  %s", r)
            elif r.severity == "WARNING":
                logger.warning("  %s", r)
            else:
                logger.info("  %s", r)


def detect_abnormal_balances(balance_rows: list) -> List[AnomalyReport]:
    """
    Fonction utilitaire pour le pipeline : détecte et journalise les soldes anormaux.

    Args:
        balance_rows: Balance N normalisée (liste de NormalizedBalanceRow)

    Returns:
        Liste d'AnomalyReport
    """
    detector = DSFAnomalyDetector()
    reports = detector.detect_abnormal_balances(balance_rows)
    detector.log_report(reports)
    return reports


__all__ = ["DSFAnomalyDetector", "AnomalyReport", "detect_abnormal_balances", "EXPECTED_NORMAL_SIDE"]
