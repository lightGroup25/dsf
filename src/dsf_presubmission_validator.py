# -*- coding: utf-8 -*-
"""
P3-MF1 : Validation pré-soumission DGI pour le DSF SYSCOHADA Cameroun.
Effectue des contrôles croisés entre le workbook DSF rempli et la balance normalisée
avant soumission à la DGI pour détecter les incohérences significatives.
"""
from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from decimal import Decimal
from typing import Dict, List, Optional

logger = logging.getLogger(__name__)

# Tolérance par défaut pour les contrôles (5% d'écart accepté)
DEFAULT_TOLERANCE_PCT = Decimal("5")  # %

# Taux CNPS patronal Cameroun (tranche A)
CNPS_RATE_PATRONAL = Decimal("11.2") / Decimal("100")


@dataclass
class ValidationResult:
    """Résultat d'un contrôle individuel de pré-soumission."""
    check_id: str
    label: str
    status: str         # "PASS", "WARN", "FAIL", "SKIP"
    detail: str = ""
    expected: Optional[str] = None
    actual: Optional[str] = None

    def __str__(self) -> str:
        parts = [f"[{self.status}] {self.check_id} — {self.label}"]
        if self.detail:
            parts.append(f"  {self.detail}")
        if self.expected is not None:
            parts.append(f"  Attendu : {self.expected}")
        if self.actual is not None:
            parts.append(f"  Obtenu  : {self.actual}")
        return "\n".join(parts)


@dataclass
class PresubmissionReport:
    """Rapport complet de pré-soumission."""
    results: List[ValidationResult] = field(default_factory=list)

    @property
    def n_pass(self)  -> int: return sum(1 for r in self.results if r.status == "PASS")
    @property
    def n_warn(self)  -> int: return sum(1 for r in self.results if r.status == "WARN")
    @property
    def n_fail(self)  -> int: return sum(1 for r in self.results if r.status == "FAIL")
    @property
    def n_skip(self)  -> int: return sum(1 for r in self.results if r.status == "SKIP")
    @property
    def is_valid(self) -> bool: return self.n_fail == 0

    @property
    def summary(self) -> str:
        return (
            f"Validation pré-soumission : {self.n_pass} PASS / "
            f"{self.n_warn} WARN / {self.n_fail} FAIL / {self.n_skip} SKIP"
        )

    def log(self) -> None:
        logger.info("=== P3-MF1 : %s ===", self.summary)
        for r in self.results:
            if r.status == "FAIL":
                logger.error("  %s", r)
            elif r.status == "WARN":
                logger.warning("  %s", r)
            elif r.status == "PASS":
                logger.info("  [PASS] %s — %s", r.check_id, r.label)


class DSFPresubmissionValidator:
    """
    P3-MF1 : Validateur pré-soumission DGI.

    Contrôles effectués :
    1. Format du NIF (M + 9 chiffres)
    2. Cohérence CA balance 70x vs CA déclaré page de garde (±5%)
    3. CNPS ≈ 11.2% masse salariale (±10%)
    4. Total actif immobilisé Bilan ≈ Total Notes 1/2/3 (±5%)
    5. Total stocks Bilan ≈ Total Note 4 (±5%)
    """

    def __init__(
        self,
        wb=None,
        balance_rows: Optional[list] = None,
        general_info=None,
        tolerance_pct: Decimal = DEFAULT_TOLERANCE_PCT,
    ):
        """
        Args:
            wb: Workbook openpyxl du DSF
            balance_rows: Balance N normalisée
            general_info: DSF_InfosGenerales (optionnel, pour NIF et CA déclaré)
            tolerance_pct: Tolérance en % pour les contrôles de cohérence
        """
        self.wb = wb
        self.balance_rows = balance_rows or []
        self.general_info = general_info
        self.tolerance = tolerance_pct
        # Index {compte: solde_net}
        self._balance_index: Dict[str, Decimal] = self._build_index(self.balance_rows)

    def generate_validation_report(self) -> PresubmissionReport:
        """Génère le rapport complet de validation pré-soumission."""
        report = PresubmissionReport()
        checks = [
            self.check_nif_format,
            self.check_ca_coherence,
            self.check_cnps_vs_masse_salariale,
            self.check_note_immo_coherence,
            self.check_note_stocks_coherence,
        ]
        for check_fn in checks:
            try:
                result = check_fn()
                report.results.append(result)
            except Exception as exc:
                logger.debug("Erreur contrôle %s : %s", check_fn.__name__, exc)
                report.results.append(ValidationResult(
                    check_id=check_fn.__name__,
                    label=check_fn.__name__,
                    status="SKIP",
                    detail=f"Erreur inattendue : {exc}",
                ))
        report.log()
        return report

    # ------------------------------------------------------------------
    # Contrôle 1 — Format NIF DGI Cameroun
    # ------------------------------------------------------------------
    def check_nif_format(self) -> ValidationResult:
        """NIF Cameroun = M + 9 chiffres. Ex: M001234567."""
        vid = "MF1-NIF"
        if self.general_info is None:
            return ValidationResult(vid, "Format NIF", "SKIP", "Pas d'info société disponible")
        nif = str(getattr(self.general_info, "nif", "") or "").strip().upper()
        if not nif:
            return ValidationResult(vid, "Format NIF", "FAIL", "NIF non renseigné", "M + 9 chiffres", "(vide)")
        if re.fullmatch(r"M\d{9}", nif):
            return ValidationResult(vid, "Format NIF", "PASS", f"NIF valide : {nif}")
        return ValidationResult(
            vid, "Format NIF", "FAIL",
            f"Format NIF incorrect : '{nif}'",
            "M + 9 chiffres (ex: M001234567)", nif,
        )

    # ------------------------------------------------------------------
    # Contrôle 2 — CA balance 70x vs CA page de garde
    # ------------------------------------------------------------------
    def check_ca_coherence(self) -> ValidationResult:
        """CA comptes 70x ≈ CA déclaré page de garde (±5%)."""
        vid = "MF1-CA"
        ca_balance = self._sum_prefix("70")

        # Lire le CA déclaré depuis la page de garde (si disponible)
        ca_declared = self._read_cell_by_label("chiffre d affaires")
        if ca_declared is None:
            return ValidationResult(
                vid, "Cohérence CA balance vs déclaré", "SKIP",
                "CA déclaré introuvable dans le workbook — contrôle ignoré",
            )

        if ca_balance == 0:
            return ValidationResult(
                vid, "Cohérence CA balance vs déclaré", "WARN",
                "CA balance nul — balance peut être incomplète",
            )

        ecart_pct = abs(ca_balance - ca_declared) / max(abs(ca_declared), Decimal("1")) * 100
        if ecart_pct <= self.tolerance:
            return ValidationResult(
                vid, "Cohérence CA balance vs déclaré", "PASS",
                f"CA balance={ca_balance:,.0f} / déclaré={ca_declared:,.0f} — écart={ecart_pct:.1f}%",
            )
        sev = "WARN" if ecart_pct <= self.tolerance * 3 else "FAIL"
        return ValidationResult(
            vid, "Cohérence CA balance vs déclaré", sev,
            f"Écart CA trop important : {ecart_pct:.1f}%",
            f"{ca_declared:,.0f} FCFA",
            f"{ca_balance:,.0f} FCFA",
        )

    # ------------------------------------------------------------------
    # Contrôle 3 — CNPS vs Masse salariale
    # ------------------------------------------------------------------
    def check_cnps_vs_masse_salariale(self) -> ValidationResult:
        """CNPS ≈ 11.2% de la masse salariale brute (tolérance ±10%)."""
        vid = "MF1-CNPS"
        masse_sal   = abs(self._sum_prefix("66"))   # Charges de personnel
        cnps_compte = abs(self._sum_prefix("431"))  # Organismes sociaux (CNPS)

        if masse_sal == 0:
            return ValidationResult(vid, "Cohérence CNPS / Masse salariale", "SKIP",
                                    "Masse salariale nulle dans la balance")

        expected_cnps = masse_sal * CNPS_RATE_PATRONAL
        if expected_cnps == 0:
            return ValidationResult(vid, "Cohérence CNPS / Masse salariale", "SKIP", "Calcul inapplicable")

        ecart_pct = abs(cnps_compte - expected_cnps) / expected_cnps * 100
        tolerance_cnps = Decimal("10")  # ±10% acceptable (cotisation multigrades)
        if ecart_pct <= tolerance_cnps:
            return ValidationResult(
                vid, "Cohérence CNPS / Masse salariale", "PASS",
                f"CNPS={cnps_compte:,.0f} / attendu≈{expected_cnps:,.0f} — écart={ecart_pct:.1f}%",
            )
        sev = "WARN" if ecart_pct <= tolerance_cnps * 2 else "FAIL"
        return ValidationResult(
            vid, "Cohérence CNPS / Masse salariale", sev,
            f"CNPS vs Masse salariale : écart {ecart_pct:.1f}% (max 10%)",
            f"≈{expected_cnps:,.0f} FCFA (11.2% de {masse_sal:,.0f})",
            f"{cnps_compte:,.0f} FCFA",
        )

    # ------------------------------------------------------------------
    # Contrôle 4 — Actif immobilisé Bilan ≈ Total Notes 1/2/3
    # ------------------------------------------------------------------
    def check_note_immo_coherence(self) -> ValidationResult:
        """Total actif immobilisé Bilan ≈ Total Notes 1+2+3 nets (±5%)."""
        vid = "MF1-NOTE-IMMO"
        # Somme comptes 2x bruts (hors 28x/29x amortissements)
        immo_brut = sum(
            (self._sum_prefix(str(p)) for p in range(20, 28)),
            Decimal("0"),
        )
        amort_immo = self._sum_prefix("28") + self._sum_prefix("29")
        immo_net_balance = abs(immo_brut) - abs(amort_immo)

        immo_net_bilan = self._read_cell_by_label("total actif immobilise")
        if immo_net_bilan is None:
            return ValidationResult(
                vid, "Cohérence Bilan actif immo vs Notes 1/2/3", "SKIP",
                "Total actif immobilisé introuvable dans le Bilan",
            )

        if immo_net_balance == 0 and immo_net_bilan == 0:
            return ValidationResult(vid, "Cohérence Bilan actif immo vs Notes 1/2/3", "PASS",
                                    "Actif immobilisé = 0 dans la balance et le bilan")

        ecart = abs(immo_net_balance - immo_net_bilan)
        base  = max(abs(immo_net_bilan), Decimal("1"))
        ecart_pct = ecart / base * 100
        if ecart_pct <= self.tolerance:
            return ValidationResult(
                vid, "Cohérence Bilan actif immo vs Notes 1/2/3", "PASS",
                f"Actif immo balance={immo_net_balance:,.0f} / bilan={immo_net_bilan:,.0f} — écart={ecart_pct:.1f}%",
            )
        sev = "WARN" if ecart_pct <= self.tolerance * 3 else "FAIL"
        return ValidationResult(
            vid, "Cohérence Bilan actif immo vs Notes 1/2/3", sev,
            f"Écart actif immobilisé : {ecart_pct:.1f}%",
            f"{immo_net_bilan:,.0f} FCFA (Bilan)",
            f"{immo_net_balance:,.0f} FCFA (balance)",
        )

    # ------------------------------------------------------------------
    # Contrôle 5 — Total stocks Bilan ≈ Total Note 4
    # ------------------------------------------------------------------
    def check_note_stocks_coherence(self) -> ValidationResult:
        """Total stocks Bilan ≈ Total Note 4 (±5%)."""
        vid = "MF1-NOTE-STOCKS"
        stocks_balance = sum(
            (self._sum_prefix(str(p)) for p in range(31, 39)),
            Decimal("0"),
        )
        stocks_bilan = self._read_cell_by_label("total stocks")
        if stocks_bilan is None:
            return ValidationResult(
                vid, "Cohérence Bilan stocks vs Note 4", "SKIP",
                "Total stocks introuvable dans le Bilan",
            )

        if stocks_balance == 0 and stocks_bilan == 0:
            return ValidationResult(vid, "Cohérence Bilan stocks vs Note 4", "PASS", "Stocks = 0")

        ecart_pct = abs(stocks_balance - stocks_bilan) / max(abs(stocks_bilan), Decimal("1")) * 100
        if ecart_pct <= self.tolerance:
            return ValidationResult(
                vid, "Cohérence Bilan stocks vs Note 4", "PASS",
                f"Stocks balance={stocks_balance:,.0f} / bilan={stocks_bilan:,.0f} — écart={ecart_pct:.1f}%",
            )
        sev = "WARN" if ecart_pct <= self.tolerance * 3 else "FAIL"
        return ValidationResult(
            vid, "Cohérence Bilan stocks vs Note 4", sev,
            f"Écart stocks : {ecart_pct:.1f}%",
            f"{stocks_bilan:,.0f} FCFA (Bilan)",
            f"{stocks_balance:,.0f} FCFA (balance)",
        )

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    def _sum_prefix(self, prefix: str) -> Decimal:
        """Somme tous les comptes de l'index dont le numéro commence par prefix."""
        return sum(
            (v for k, v in self._balance_index.items() if k.startswith(prefix)),
            Decimal("0"),
        )

    @staticmethod
    def _build_index(rows: list) -> Dict[str, Decimal]:
        """Construit {compte: solde_net} depuis NormalizedBalanceRow."""
        index: Dict[str, Decimal] = {}
        for row in rows:
            try:
                compte  = str(getattr(row, "compte", "") or "").strip()
                closing = getattr(row, "closing_balance", None)
                if closing is not None:
                    solde = Decimal(str(closing))
                else:
                    debit  = Decimal(str(getattr(row, "debit_total",  0) or 0))
                    credit = Decimal(str(getattr(row, "credit_total", 0) or 0))
                    solde  = debit - credit
                if compte:
                    index[compte] = index.get(compte, Decimal("0")) + solde
            except Exception:
                pass
        return index

    def _read_cell_by_label(self, label_norm_fragment: str) -> Optional[Decimal]:
        """
        Cherche dans toutes les feuilles du workbook une ligne dont le libellé
        contient `label_norm_fragment` (texte normalisé), et lit la valeur numérique
        la plus à droite sur cette ligne.
        """
        if self.wb is None:
            return None
        import unicodedata, re as _re

        def _norm(s: str) -> str:
            s = unicodedata.normalize("NFD", s)
            s = "".join(c for c in s if unicodedata.category(c) != "Mn")
            return _re.sub(r"\s+", " ", s.lower()).strip()

        fragment = _norm(label_norm_fragment)
        for ws in self.wb.worksheets:
            for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
                for cell in row[:4]:
                    if isinstance(cell.value, str) and fragment in _norm(cell.value):
                        # Chercher la valeur numérique la plus à droite sur cette ligne
                        for c in reversed(row):
                            if isinstance(c.value, (int, float)) and c.value != 0:
                                return Decimal(str(c.value))
        return None


# ---------------------------------------------------------------------------
# Fonction utilitaire pour le pipeline
# ---------------------------------------------------------------------------
def run_presubmission_validation(
    wb,
    balance_rows: list,
    general_info=None,
) -> PresubmissionReport:
    """
    Exécute la validation pré-soumission DGI et retourne le rapport.
    Intégration dans dsf_pipeline.py après run_fiscal_controls().
    """
    validator = DSFPresubmissionValidator(wb, balance_rows, general_info)
    report = validator.generate_validation_report()
    return report


__all__ = [
    "DSFPresubmissionValidator",
    "PresubmissionReport",
    "ValidationResult",
    "run_presubmission_validation",
]
