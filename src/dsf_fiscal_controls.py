# -*- coding: utf-8 -*-
"""
Module de contrôles fiscaux DSF — Priorité 2 (A1)
Contrôles de cohérence : TVA, IS, CR=Bilan, Capitaux propres.
Complète dsf_controls.py sans le modifier (rétrocompatibilité).
"""
from __future__ import annotations

import logging
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Dict, List, Optional

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _dec(value: Any) -> Decimal:
    """Convertit une valeur en Decimal, retourne 0 si invalide."""
    if value is None:
        return Decimal("0")
    if isinstance(value, Decimal):
        return value
    try:
        return Decimal(str(value))
    except (InvalidOperation, ValueError, TypeError):
        return Decimal("0")


def _read_cell(ws, ref: str) -> Decimal:
    """Lit une cellule du workbook et retourne sa valeur en Decimal."""
    try:
        val = ws[ref].value
        return _dec(val)
    except Exception:
        return Decimal("0")


# ---------------------------------------------------------------------------
# Dataclasses
# ---------------------------------------------------------------------------

@dataclass
class FiscalIssue:
    name: str
    status: str          # PASS | WARN | FAIL
    message: str
    details: Dict[str, Any] = field(default_factory=dict)


@dataclass
class FiscalReport:
    issues: List[FiscalIssue]

    @property
    def summary(self) -> Dict[str, int]:
        counts = {"PASS": 0, "WARN": 0, "FAIL": 0}
        for issue in self.issues:
            if issue.status in counts:
                counts[issue.status] += 1
        return counts

    def has_failures(self) -> bool:
        return self.summary.get("FAIL", 0) > 0

    def to_dict(self) -> Dict[str, Any]:
        return {
            "summary": self.summary,
            "issues": [
                {
                    "name": i.name,
                    "status": i.status,
                    "message": i.message,
                    "details": i.details,
                }
                for i in self.issues
            ],
        }


# ---------------------------------------------------------------------------
# Référentiel des cellules DSF par contrôle
# Ces références correspondent au template "DSF Normal Standard.xlsx"
# ---------------------------------------------------------------------------

# Bilan Paysage — Actif
BILAN_TOTAL_ACTIF_N   = "D99"   # Total Actif N   (ligne totale Actif)
BILAN_TOTAL_ACTIF_N1  = "F99"   # Total Actif N-1

# Bilan Paysage — Passif
BILAN_TOTAL_PASSIF_N  = "K99"   # Total Passif N
BILAN_TOTAL_PASSIF_N1 = "M99"   # Total Passif N-1

# Compte de Résultat
CR_RESULTAT_NET_N     = "D99"   # Résultat net N  (feuille COMPTE DE RESULTAT)
CR_RESULTAT_NET_N1    = "F99"   # Résultat net N-1

# Bilan — Résultat de l'exercice (capitaux propres)
BILAN_RESULTAT_N      = "K13"   # Résultat exercice au Passif (capitaux propres)

# TVA — feuille COMPTE DE RESULTAT ou annexe TVA
TVA_COLLECTEE_CELL    = "D45"   # TVA collectée (produits)
TVA_DEDUCTIBLE_CELL   = "D30"   # TVA déductible (charges)
TVA_NETTE_CELL        = "D46"   # TVA nette à payer

# IS — Impôt sur les sociétés
IS_CELL               = "D50"   # IS comptabilisé
RESULTAT_AVANT_IS     = "D49"   # Résultat avant IS

# Capitaux propres
CAPITAUX_PROPRES_CELL = "K20"   # Total capitaux propres
CAPITAL_CELL          = "K10"   # Capital social
RESERVES_CELL         = "K11"   # Réserves
REPORT_NOUVEAU_CELL   = "K12"   # Report à nouveau


# ---------------------------------------------------------------------------
# Suite de contrôles fiscaux
# ---------------------------------------------------------------------------

class FiscalControlSuite:
    """
    Contrôles fiscaux DGI appliqués sur le workbook DSF généré.
    Lit directement les cellules du fichier Excel produit.
    """

    TOLERANCE = Decimal("1000")   # Tolérance d'arrondi en FCFA (1 000 FCFA)

    def __init__(self, tolerance: Decimal = TOLERANCE):
        self.tolerance = tolerance

    def evaluate_from_workbook(self, dsf_path: Path) -> FiscalReport:
        """
        Charge le DSF généré et exécute tous les contrôles fiscaux.

        Args:
            dsf_path: Chemin vers le fichier DSF généré (.xlsx)

        Returns:
            FiscalReport avec la liste des issues
        """
        try:
            from openpyxl import load_workbook
            wb = load_workbook(dsf_path, data_only=True)
        except Exception as exc:
            logger.error("Impossible de charger le DSF pour les contrôles fiscaux : %s", exc)
            return FiscalReport(issues=[
                FiscalIssue(
                    name="load_error",
                    status="FAIL",
                    message=f"Impossible de charger le DSF : {exc}",
                )
            ])

        issues: List[FiscalIssue] = []
        issues.append(self._check_bilan_equilibre(wb))
        issues.append(self._check_cr_bilan_coherence(wb))
        issues.append(self._check_tva_coherence(wb))
        issues.append(self._check_is_coherence(wb))
        issues.append(self._check_capitaux_propres(wb))

        wb.close()
        return FiscalReport(issues=issues)

    # ------------------------------------------------------------------
    # Contrôles individuels
    # ------------------------------------------------------------------

    def _check_bilan_equilibre(self, wb) -> FiscalIssue:
        """Actif total N == Passif total N (équation bilancielle)."""
        ws = self._get_sheet(wb, "BILAN PAYSAGE", "BILAN")
        if ws is None:
            return FiscalIssue(
                name="bilan_equilibre",
                status="WARN",
                message="Feuille BILAN PAYSAGE introuvable — contrôle ignoré",
            )

        actif  = _read_cell(ws, BILAN_TOTAL_ACTIF_N)
        passif = _read_cell(ws, BILAN_TOTAL_PASSIF_N)
        diff   = abs(actif - passif)
        status = "PASS" if diff <= self.tolerance else "FAIL"
        return FiscalIssue(
            name="bilan_equilibre",
            status=status,
            message="Actif = Passif ✓" if status == "PASS" else f"Déséquilibre Actif/Passif : {diff:,.0f} FCFA",
            details={"actif_N": str(actif), "passif_N": str(passif), "ecart": str(diff)},
        )

    def _check_cr_bilan_coherence(self, wb) -> FiscalIssue:
        """
        Résultat net du Compte de Résultat == Résultat de l'exercice au Bilan.
        Contrôle fondamental exigé par la DGI.
        """
        ws_cr    = self._get_sheet(wb, "COMPTE DE RESULTAT", "CR", "RESULTAT")
        ws_bilan = self._get_sheet(wb, "BILAN PAYSAGE", "BILAN")

        if ws_cr is None or ws_bilan is None:
            return FiscalIssue(
                name="cr_bilan_coherence",
                status="WARN",
                message="Feuilles CR ou Bilan introuvables — contrôle ignoré",
            )

        resultat_cr    = _read_cell(ws_cr,    CR_RESULTAT_NET_N)
        resultat_bilan = _read_cell(ws_bilan, BILAN_RESULTAT_N)
        diff = abs(resultat_cr - resultat_bilan)
        status = "PASS" if diff <= self.tolerance else "FAIL"
        return FiscalIssue(
            name="cr_bilan_coherence",
            status=status,
            message="Résultat CR = Résultat Bilan ✓" if status == "PASS"
                    else f"Résultat CR ≠ Résultat Bilan : écart {diff:,.0f} FCFA",
            details={
                "resultat_CR": str(resultat_cr),
                "resultat_Bilan": str(resultat_bilan),
                "ecart": str(diff),
            },
        )

    def _check_tva_coherence(self, wb) -> FiscalIssue:
        """
        TVA collectée - TVA déductible ≈ TVA nette déclarée.
        Contrôle de cohérence TVA exigé par la DGI.
        """
        ws = self._get_sheet(wb, "COMPTE DE RESULTAT", "CR", "RESULTAT")
        if ws is None:
            return FiscalIssue(
                name="tva_coherence",
                status="WARN",
                message="Feuille CR introuvable — contrôle TVA ignoré",
            )

        tva_collectee  = _read_cell(ws, TVA_COLLECTEE_CELL)
        tva_deductible = _read_cell(ws, TVA_DEDUCTIBLE_CELL)
        tva_nette      = _read_cell(ws, TVA_NETTE_CELL)

        if tva_collectee == 0 and tva_deductible == 0:
            return FiscalIssue(
                name="tva_coherence",
                status="WARN",
                message="TVA collectée et déductible à zéro — vérifier le remplissage",
                details={"tva_collectee": "0", "tva_deductible": "0"},
            )

        tva_calculee = tva_collectee - tva_deductible
        diff = abs(tva_calculee - tva_nette)
        status = "PASS" if diff <= self.tolerance else "FAIL"
        return FiscalIssue(
            name="tva_coherence",
            status=status,
            message="TVA nette cohérente ✓" if status == "PASS"
                    else f"TVA nette incohérente : écart {diff:,.0f} FCFA",
            details={
                "tva_collectee": str(tva_collectee),
                "tva_deductible": str(tva_deductible),
                "tva_nette_calculee": str(tva_calculee),
                "tva_nette_declaree": str(tva_nette),
                "ecart": str(diff),
            },
        )

    def _check_is_coherence(self, wb) -> FiscalIssue:
        """
        IS comptabilisé cohérent avec le résultat avant IS.
        Taux IS Cameroun : 33% (grandes entreprises).
        """
        ws = self._get_sheet(wb, "COMPTE DE RESULTAT", "CR", "RESULTAT")
        if ws is None:
            return FiscalIssue(
                name="is_coherence",
                status="WARN",
                message="Feuille CR introuvable — contrôle IS ignoré",
            )

        resultat_avant_is = _read_cell(ws, RESULTAT_AVANT_IS)
        is_comptabilise   = _read_cell(ws, IS_CELL)

        if resultat_avant_is <= 0:
            return FiscalIssue(
                name="is_coherence",
                status="WARN",
                message="Résultat avant IS ≤ 0 — IS non applicable ou déficit",
                details={"resultat_avant_is": str(resultat_avant_is)},
            )

        if is_comptabilise == 0:
            return FiscalIssue(
                name="is_coherence",
                status="WARN",
                message="IS à zéro alors que le résultat avant IS est positif — vérifier",
                details={"resultat_avant_is": str(resultat_avant_is)},
            )

        # Taux IS Cameroun : 33% (grandes entreprises DGE)
        taux_is = Decimal("0.33")
        is_theorique = (resultat_avant_is * taux_is).quantize(Decimal("1"))
        diff = abs(is_comptabilise - is_theorique)
        # Tolérance élargie pour IS (différences de base imposable, crédits d'impôt)
        tolerance_is = max(self.tolerance, resultat_avant_is * Decimal("0.05"))
        status = "PASS" if diff <= tolerance_is else "WARN"
        return FiscalIssue(
            name="is_coherence",
            status=status,
            message="IS cohérent avec résultat avant IS ✓" if status == "PASS"
                    else f"IS potentiellement incohérent (écart {diff:,.0f} FCFA vs théorique {is_theorique:,.0f})",
            details={
                "resultat_avant_is": str(resultat_avant_is),
                "is_comptabilise": str(is_comptabilise),
                "is_theorique_33pct": str(is_theorique),
                "ecart": str(diff),
            },
        )

    def _check_capitaux_propres(self, wb) -> FiscalIssue:
        """
        Capitaux propres = Capital + Réserves + Report à nouveau + Résultat.
        Contrôle de cohérence des fonds propres.
        """
        ws = self._get_sheet(wb, "BILAN PAYSAGE", "BILAN")
        if ws is None:
            return FiscalIssue(
                name="capitaux_propres",
                status="WARN",
                message="Feuille Bilan introuvable — contrôle capitaux propres ignoré",
            )

        capital        = _read_cell(ws, CAPITAL_CELL)
        reserves       = _read_cell(ws, RESERVES_CELL)
        report_nouveau = _read_cell(ws, REPORT_NOUVEAU_CELL)
        resultat       = _read_cell(ws, BILAN_RESULTAT_N)
        cp_total       = _read_cell(ws, CAPITAUX_PROPRES_CELL)

        cp_calcule = capital + reserves + report_nouveau + resultat
        diff = abs(cp_calcule - cp_total)
        status = "PASS" if diff <= self.tolerance else "FAIL"
        return FiscalIssue(
            name="capitaux_propres",
            status=status,
            message="Capitaux propres cohérents ✓" if status == "PASS"
                    else f"Capitaux propres incohérents : écart {diff:,.0f} FCFA",
            details={
                "capital": str(capital),
                "reserves": str(reserves),
                "report_nouveau": str(report_nouveau),
                "resultat": str(resultat),
                "cp_calcule": str(cp_calcule),
                "cp_total_bilan": str(cp_total),
                "ecart": str(diff),
            },
        )

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _get_sheet(wb, *candidates):
        """Retourne la première feuille dont le nom contient l'un des candidats."""
        for name in wb.sheetnames:
            name_up = name.upper()
            for c in candidates:
                if c.upper() in name_up:
                    return wb[name]
        return None


# ---------------------------------------------------------------------------
# Intégration pipeline : fonction utilitaire
# ---------------------------------------------------------------------------

def run_fiscal_controls(dsf_path: Path, tolerance: Decimal = FiscalControlSuite.TOLERANCE) -> FiscalReport:
    """
    Lance les contrôles fiscaux sur un DSF généré.

    Args:
        dsf_path: Chemin vers le fichier DSF .xlsx
        tolerance: Tolérance d'arrondi en FCFA

    Returns:
        FiscalReport
    """
    suite = FiscalControlSuite(tolerance=tolerance)
    report = suite.evaluate_from_workbook(dsf_path)
    summary = report.summary
    logger.info(
        "Contrôles fiscaux : %d PASS / %d WARN / %d FAIL",
        summary["PASS"], summary["WARN"], summary["FAIL"],
    )
    for issue in report.issues:
        level = logging.ERROR if issue.status == "FAIL" else (
            logging.WARNING if issue.status == "WARN" else logging.INFO
        )
        logger.log(level, "[%s] %s — %s", issue.status, issue.name, issue.message)
    return report


__all__ = ["FiscalControlSuite", "FiscalIssue", "FiscalReport", "run_fiscal_controls"]
