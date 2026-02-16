# -*- coding: utf-8 -*-
"""
Semantic Balance Filler - Intelligent cell-by-cell matching
Analyzes cell labels (row + column), fuzzy-matches balance accounts, and fills values.
"""
from __future__ import annotations

import json
import logging
import re
import unicodedata
from collections import Counter
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from difflib import SequenceMatcher
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.utils import coordinate_to_tuple, get_column_letter

from balance_normalizer import BalanceNormalizer, NormalizedBalanceRow
from dsf_inventory import DSFInventory
from column_formula_detector import ColumnFormulaDetector

logger = logging.getLogger(__name__)


@dataclass
class CellNeed:
    """Represents a need from a single cell"""
    sheet: str
    cell: str
    row_num: int
    col_num: int
    row_label: Optional[str] = None
    col_label: Optional[str] = None
    is_merged: bool = False
    merge_range: Optional[str] = None


@dataclass
class MatchedAccount:
    """Single account match for a cell"""
    compte: str
    label: str
    amount: Decimal
    side: str  # "debit" or "credit"
    similarity_score: float


@dataclass
class CellAssignment:
    """Final assignment for a cell"""
    sheet: str
    cell: str
    row_label: str
    col_label: str
    is_total: bool
    matched_accounts: List[MatchedAccount] = field(default_factory=list)
    total_amount: Decimal = Decimal("0")
    source_accounts: List[str] = field(default_factory=list)
    confidence: float = 0.0
    notes: str = ""


@dataclass(frozen=True)
class BusinessRule:
    sheet_groups: Tuple[str, ...]
    label_keywords: Tuple[str, ...]
    prefixes: Tuple[str, ...]
    classes: Tuple[str, ...] = ()
    aggregate: bool = False
    min_score: float = 0.60
    max_accounts: int = 12
    allow_n1_fallback: bool = False
    value_source_override: Optional[str] = None
    allow_stock_fallback_when_zero: bool = False


class SemanticBalanceFiller:
    def __init__(
        self,
        template_path: Path | str,
        balance_file: Path | str,
        inventory: DSFInventory,
        fuzzy_threshold: float = 0.70,
        enable_hybrid: bool = False,
        column_overrides: Optional[Dict[str, object]] = None,
    ):
        self.template_path = Path(template_path)
        self.balance_file = Path(balance_file)
        self.inventory = inventory
        self.fuzzy_threshold = fuzzy_threshold
        self.min_match_score = max(0.62, fuzzy_threshold)
        self.enable_hybrid = enable_hybrid
        self.column_overrides = column_overrides or {
            "compte": 1,
            "label": 4,
            "debit_columns": [24, 21, 17, 11],
            "credit_columns": [27, 26, 20, 14, 13],
        }
        self.wb = None
        self.balance_accounts: Dict[str, NormalizedBalanceRow] = {}
        self.assignments: List[CellAssignment] = []
        self.unmatched_cells: List[CellNeed] = []
        self.formula_detectors: Dict[str, ColumnFormulaDetector] = {}  # {sheet_name: detector}
        self.formulas_applied: int = 0
        self.formula_logs: List[Dict[str, str]] = []
        self.account_usage: Counter[str] = Counter()
        self.business_rules: List[BusinessRule] = self._build_business_rules()
        self.account_time_slices: Dict[str, Dict[str, Decimal]] = {}

    def load(self) -> None:
        """Load template and normalize balance"""
        if not self.template_path.exists():
            raise FileNotFoundError(f"Template not found: {self.template_path}")
        if not self.balance_file.exists():
            raise FileNotFoundError(f"Balance not found: {self.balance_file}")

        # Load template workbook before scanning formulas.
        self.wb = load_workbook(self.template_path, data_only=False)

        # Detect formulas in column headers for all sheets.
        logger.info("Detecting column formulas in template...")
        for sheet_name in self.wb.sheetnames:
            detector = ColumnFormulaDetector(self.wb[sheet_name])
            formulas = detector.detect()
            if formulas:
                self.formula_detectors[sheet_name] = detector
                logger.info(f"  {sheet_name}: {len(formulas)} formula columns detected")
        
        normalizer = BalanceNormalizer(self.balance_file, column_overrides=self.column_overrides)
        for normalized_row in normalizer.iterate():
            self.balance_accounts[normalized_row.compte] = normalized_row
        logger.info(f"Loaded {len(self.balance_accounts)} accounts from balance")

        # Extract opening/movement/closing slices for N / N-1 aware mapping.
        self.account_time_slices = self._extract_balance_time_slices()
        logger.info("Loaded timeslices for %s accounts", len(self.account_time_slices))

        non_zero = sum(1 for row in self.balance_accounts.values() if row.debit_balance > 0 or row.credit_balance > 0)
        logger.info(f"  Accounts with non-zero amounts: {non_zero}/{len(self.balance_accounts)}")

    def fill(self) -> int:
        """
        Main filling logic: iterate sheets, rows, cells
        Also detects and skips columns with formulas (will be applied later)
        Returns number of assignments made
        """
        if self.wb is None or not self.balance_accounts:
            raise RuntimeError("Filler not loaded. Call load() first.")

        assignments_count = 0

        EXCLUDED_SHEETS = ["ENTETE", "ENTÊTE", "Fiche R1", "R1", "Fiche R2", "R2", "Fiche R3", "R3"]

        for sheet_name in self.wb.sheetnames:
            if sheet_name in EXCLUDED_SHEETS:
                logger.info(f"Skipping general info sheet: {sheet_name}")
                continue

            ws = self.wb[sheet_name]
            logger.info(f"Processing sheet: {sheet_name}")

            has_formula_detector = sheet_name in self.formula_detectors
            formula_detector = self.formula_detectors.get(sheet_name) if has_formula_detector else None

            merged_coords = set()
            for merged_range in ws.merged_cells.ranges:
                for cell in merged_range.cells:
                    merged_coords.add(f"{cell[0]}{cell[1]}")

            for row_num in range(11, ws.max_row + 1):
                row_label = self._extract_row_label(ws, row_num)
                if not row_label:
                    continue

                for col_num in range(1, ws.max_column + 1):
                    cell = ws.cell(row=row_num, column=col_num)

                    if cell.coordinate in merged_coords:
                        continue

                    col_letter = get_column_letter(col_num)
                    if formula_detector and formula_detector.has_formula(col_letter):
                        continue

                    col_label = self._extract_col_label(ws, col_num)
                    field = self.inventory.get_field(sheet_name, cell.coordinate) if self.inventory else None

                    if not self._is_inventory_writable(sheet_name, cell.coordinate, col_label):
                        continue

                    if not self._is_fillable_cell(cell):
                        continue

                    if not col_label:
                        continue

                    if not self._is_supported_numeric_target(field, row_label, col_label):
                        continue

                    need = CellNeed(
                        sheet=sheet_name,
                        cell=cell.coordinate,
                        row_num=row_num,
                        col_num=col_num,
                        row_label=row_label,
                        col_label=col_label,
                        is_merged=False,
                    )

                    assignment = self._satisfy_cell_need(need, field)
                    if assignment:
                        self.assignments.append(assignment)
                        self._write_cell(ws, need, assignment)
                        self.account_usage.update(assignment.source_accounts)
                        assignments_count += 1
                    else:
                        self.unmatched_cells.append(need)

        logger.info("Applying column formulas...")
        self._apply_column_formulas()
        logger.info(f"Applied {self.formulas_applied} column formulas")
        
        return assignments_count

    def _is_inventory_writable(self, sheet: str, cell_ref: str, col_label: Optional[str]) -> bool:
        """Strict validation against inventory: locked/formula/type expected."""
        if not self.inventory:
            return True
        field = self.inventory.get_field(sheet, cell_ref)
        if not field:
            return False
        if field.has_formula:
            return False

        if field.locked:
            sheet_upper = sheet.upper()
            if "NOTE" not in sheet_upper and "SYNTHESE" not in sheet_upper and "SYNTH" not in sheet_upper:
                return False

        required_type = self._infer_required_column_type(col_label)
        if required_type and field.column_type and field.column_type != required_type:
            return False
        return True

    def _infer_required_column_type(self, col_label: Optional[str]) -> Optional[str]:
        """Infer expected inventory column_type from column label (strict)."""
        if not col_label:
            return None
        label = col_label.lower()
        if "n-1" in label or "n - 1" in label or "n- 1" in label:
            return "exercice_n1"
        if ("exercice" in label or "exerc" in label) and "n-1" not in label and "n - 1" not in label:
            return "exercice_n"
        return None

    def _is_supported_numeric_target(self, field, row_label: str, col_label: str) -> bool:
        """Reject structural/text columns and keep numeric targets only."""
        if not row_label or not col_label:
            return False

        col_letter = getattr(field, "column_letter", "")
        if col_letter in {"A", "B", "C"}:
            # Most templates use these columns for codes/labels/sign/note refs.
            return False

        sheet = getattr(field, "sheet", "") if field else ""
        row_norm_full = self._normalize_text(row_label)
        if sheet.upper() == "TABLEAU DES FLUX DE TRESORERIE":
            # Only value columns
            if col_letter not in {"E", "F"}:
                return False
            # Keep leaf operational rows; skip section headers/totals/controls/footnotes.
            flux_forbidden = (
                "flux de tresorerie provenant",
                "somme",
                "variation de la tresorerie nette",
                "tresorerie nette au 31 decembre",
                "controle tresorerie",
            )
            if any(k in row_norm_full for k in flux_forbidden):
                return False
            if row_norm_full.startswith("1 a l exclusion") or row_norm_full.startswith("1 a l"):
                return False

        if sheet.upper().strip() == "NOTE 13":
            if col_letter != "F":
                return False
            note13_forbidden = ("commentaire", "indiquer", "avantages", "delai restant")
            if any(k in row_norm_full for k in note13_forbidden):
                return False

        row_clean = self._normalize_text(row_label)
        col_clean = self._normalize_text(col_label)
        if self._is_structural_text(row_clean) or self._is_structural_text(col_clean):
            return False

        combined = f"{row_clean} {col_clean}"
        forbidden_keywords = {
            "reference", "ref", "libelle", "rubrique", "signe", "code", "notes", "note",
            "designation", "identification", "duree", "unite", "tableau",
        }
        if any(kw in combined for kw in forbidden_keywords):
            return False

        # Ratios/percentages are usually computed columns, not direct balance mapping.
        if "%" in col_label or any(kw in col_clean for kw in ("ratio", "taux", "pourcentage")):
            return False

        numeric_hints = (
            "exercice", "montant", "solde", "net", "brut", "debit", "credit", "cloture",
            "ouverture", "variation", "flux", "acquisition", "cession", "virement", "apport",
        )
        if any(kw in combined for kw in numeric_hints):
            return True

        if field:
            fmt = str(getattr(field, "number_format", "") or "").replace(" ", "")
            if any(ch in fmt for ch in ("0", "#")):
                return True
            validations = getattr(field, "data_validation", []) or []
            if any((v.get("type") in ("decimal", "whole")) for v in validations if isinstance(v, dict)):
                return True
        return False

    def _extract_row_label(self, ws, row_num: int) -> Optional[str]:
        """Extract semantic row label, preferring description columns over code columns."""
        col_a = ws.cell(row=row_num, column=1).value
        col_b = ws.cell(row=row_num, column=2).value

        b_text = col_b.strip() if isinstance(col_b, str) else ""
        if b_text and not self._is_structural_text(self._normalize_text(b_text)):
            return b_text

        a_text = col_a.strip() if isinstance(col_a, str) else ""
        if a_text and not re.fullmatch(r"[A-Z0-9]{1,4}", a_text):
            return a_text

        if b_text:
            return b_text
        if a_text:
            return a_text
        return None

    def _extract_col_label(self, ws, col_num: int) -> Optional[str]:
        """Build a stable column label from header rows."""
        parts: List[str] = []
        seen: Set[str] = set()

        for row in range(1, 13):
            cell = ws.cell(row=row, column=col_num)
            if not isinstance(cell.value, str):
                continue
            val = " ".join(cell.value.strip().split())
            if not val:
                continue
            val_norm = self._normalize_text(val)
            if any(
                noise in val_norm
                for noise in ("designation entite", "numero d identification", "duree", "au 31 decembre")
            ):
                continue
            if len(val_norm) > 120:
                continue
            if val not in seen:
                seen.add(val)
                parts.append(val)

        if not parts:
            return None

        priority = [p for p in parts if any(k in self._normalize_text(p) for k in ("exercice", "montant", "solde", "net", "brut", "debit", "credit", "ouverture", "cloture"))]
        if priority:
            return " | ".join(priority[:2])
        return " | ".join(parts[:2])

    def _is_fillable_cell(self, cell) -> bool:
        """Check if cell is fillable (empty, no formula, structural)"""
        if cell.value is not None:
            return False
        if cell.data_type == "f":  # Formula
            return False
        # Ignore merged cells for now (read-only in openpyxl)
        if cell.data_type == "e":  # Empty merged cell
            return False
        return True

    def _satisfy_cell_need(self, need: CellNeed, field=None) -> Optional[CellAssignment]:
        """
        Strict cell-by-cell logic:
        - infer expected data source from column semantics (N/N-1, opening/closing/movement)
        - match account labels mainly on row semantics
        - enforce confidence and ambiguity checks
        """
        if not need.row_label:
            return None

        row_label_lower = self._normalize_text(need.row_label)
        is_total = (
            "total" in row_label_lower
            or "somme" in row_label_lower
            or self._looks_like_section_heading(need.row_label)
        )
        col_type = self._extract_column_type(need.col_label)
        year_hint = self._infer_year_hint(need.col_label)
        value_source = self._infer_value_source(need.col_label)
        effective_source = "opening" if year_hint == "n1" and value_source == "final" else value_source
        expected_classes = self._expected_classes_for_sheet(need.sheet, need.cell)
        preferred_prefixes = self._preferred_prefixes_for_label(need.row_label)

        matches = self._fuzzy_match_accounts(
            row_label=need.row_label,
            col_type=col_type,
            is_total=is_total,
            year_hint=year_hint,
            value_source=value_source,
            expected_classes=expected_classes,
            preferred_prefixes=preferred_prefixes,
            sheet_name=need.sheet,
            cell_ref=need.cell,
        )
        if not matches:
            matches = self._match_with_business_rules(
                row_label=need.row_label,
                sheet_name=need.sheet,
                cell_ref=need.cell,
                col_type=col_type,
                year_hint=year_hint,
                value_source=value_source,
                expected_classes=expected_classes,
            )

        if not matches:
            logger.debug(f"No matches for {need.cell}: row+col composite")
            return None

        total_amount = sum(m.amount for m in matches)
        if total_amount <= 0:
            return None

        assignment = CellAssignment(
            sheet=need.sheet,
            cell=need.cell,
            row_label=need.row_label,
            col_label=need.col_label,
            is_total=is_total,
            matched_accounts=matches,
            total_amount=total_amount,
            source_accounts=[m.compte for m in matches],
            confidence=matches[0].similarity_score if matches else 0.0,
            notes=f"source={effective_source}; year={year_hint or 'n'}",
        )

        return assignment

    def _fuzzy_match_accounts(
        self,
        row_label: str,
        col_type: Optional[str],
        is_total: bool,
        year_hint: Optional[str] = None,
        value_source: str = "final",
        expected_classes: Optional[Set[str]] = None,
        preferred_prefixes: Optional[Set[str]] = None,
        sheet_name: Optional[str] = None,
        cell_ref: Optional[str] = None,
    ) -> List[MatchedAccount]:
        matches: List[Tuple[float, MatchedAccount]] = []
        row_tokens = self._tokenize(row_label)
        row_norm = self._normalize_text(row_label)

        for compte, normalized_row in self.balance_accounts.items():
            if expected_classes and normalized_row.classe not in expected_classes:
                continue
            if sheet_name and cell_ref and not self._sheet_cell_domain_filter(sheet_name, cell_ref, normalized_row):
                continue

            amount = self._resolve_amount(normalized_row, col_type, year_hint, value_source)
            if amount <= 0:
                continue

            score = self._score_account_match(row_label, row_tokens, normalized_row, row_norm=row_norm)
            score -= min(0.18, self.account_usage.get(compte, 0) * 0.02)
            if preferred_prefixes and not any(compte.startswith(prefix) for prefix in preferred_prefixes):
                score -= 0.10

            if score < self.min_match_score:
                continue

            matched = MatchedAccount(
                compte=compte,
                label=normalized_row.label,
                amount=amount,
                side=col_type or normalized_row.natural_side,
                similarity_score=score,
            )
            matches.append((score, matched))

        matches.sort(key=lambda x: (-x[0], x[1].compte))
        if not matches:
            return []

        # Ambiguity guard: if top candidates are too close, keep unmatched.
        if not is_total and len(matches) > 1:
            gap = matches[0][0] - matches[1][0]
            if gap < 0.07 and matches[0][0] < 0.86:
                return []

        if is_total:
            best_score = matches[0][0]
            accepted = [m for s, m in matches if s >= max(self.min_match_score, best_score - 0.08)]
            return accepted[:12]
        return [matches[0][1]]

    def _score_account_match(
        self,
        row_label: str,
        row_tokens: Set[str],
        normalized_row: NormalizedBalanceRow,
        row_norm: Optional[str] = None,
    ) -> float:
        account_label = normalized_row.label
        if row_norm is None:
            row_norm = self._normalize_text(row_label)
        acc_norm = self._normalize_text(account_label)
        fuzzy = self._fuzzy_similarity(row_norm, acc_norm)
        account_tokens = self._tokenize(account_label)
        overlap = self._token_overlap(row_tokens, account_tokens)
        score = (0.58 * fuzzy) + (0.42 * overlap)

        if "incorporelle" in row_norm and "corporelle" in acc_norm and "incorporelle" not in acc_norm:
            score -= 0.20
        if "corporelle" in row_norm and "incorporelle" in acc_norm and "corporelle" not in acc_norm:
            score -= 0.20
        return score

    def _resolve_amount(
        self,
        normalized_row: NormalizedBalanceRow,
        col_type: Optional[str],
        year_hint: Optional[str],
        value_source: str,
    ) -> Decimal:
        raw = normalized_row.metadata.get("raw", {}) if isinstance(normalized_row.metadata, dict) else {}
        slices = self.account_time_slices.get(normalized_row.compte, {})

        opening_debit = slices.get("opening_debit", Decimal("0"))
        opening_credit = slices.get("opening_credit", Decimal("0"))
        movement_debit = slices.get("movement_debit", self._as_decimal(raw.get("debit")))
        movement_credit = slices.get("movement_credit", self._as_decimal(raw.get("credit")))
        closing_debit = slices.get("closing_debit", normalized_row.debit_balance)
        closing_credit = slices.get("closing_credit", normalized_row.credit_balance)
        opening_signed = slices.get("opening_signed", self._as_decimal(raw.get("solde_initial")))
        closing_signed = slices.get("closing_signed", normalized_row.solde_final)
        flux_value = slices.get("flux", self._as_decimal(raw.get("flux_tresorerie")))

        source = value_source
        if year_hint == "n1" and source == "final":
            source = "opening"

        if source == "opening":
            signed = opening_signed
        elif source == "movement_debit":
            signed = movement_debit
        elif source == "movement_credit":
            signed = -movement_credit
        elif source == "variation":
            signed = closing_signed - opening_signed
        elif source == "flux":
            signed = flux_value
        else:
            signed = closing_signed

        if source in {"movement_debit", "movement_credit"}:
            if col_type == "debit":
                return movement_debit if source == "movement_debit" else Decimal("0")
            if col_type == "credit":
                return movement_credit if source == "movement_credit" else Decimal("0")
            return movement_debit if source == "movement_debit" else movement_credit

        if source == "opening":
            if col_type == "debit":
                return opening_debit
            if col_type == "credit":
                return opening_credit
            return signed.copy_abs() if signed != 0 else Decimal("0")

        if source == "final":
            if col_type == "debit":
                return closing_debit
            if col_type == "credit":
                return closing_credit
            return signed.copy_abs() if signed != 0 else Decimal("0")

        if col_type == "debit":
            return signed if signed > 0 else Decimal("0")
        if col_type == "credit":
            return -signed if signed < 0 else Decimal("0")

        # No debit/credit hint: keep absolute positive amount.
        if signed == 0:
            return Decimal("0")
        return signed.copy_abs()

    def _extract_balance_time_slices(self) -> Dict[str, Dict[str, Decimal]]:
        """Read the balance sheet once to expose opening/movement/closing values by account."""
        wb = load_workbook(self.balance_file, data_only=True)
        try:
            ws = wb.active
            compte_col = int(self.column_overrides.get("compte", 1))
            label_col = int(self.column_overrides.get("label", 4))

            # Prefer explicit overrides, fallback to conventional GULFCAM positions.
            opening_debit_cols = self._resolve_override_columns("opening_debit_columns", [11])
            opening_credit_cols = self._resolve_override_columns("opening_credit_columns", [14, 13])
            movement_debit_cols = self._resolve_override_columns("movement_debit_columns", [17])
            movement_credit_cols = self._resolve_override_columns("movement_credit_columns", [20])
            closing_debit_cols = self._resolve_override_columns("closing_debit_columns", [24, 21])
            closing_credit_cols = self._resolve_override_columns("closing_credit_columns", [27, 26])

            slices: Dict[str, Dict[str, Decimal]] = {}
            for row_idx in range(13, ws.max_row + 1):
                compte_raw = ws.cell(row=row_idx, column=compte_col).value
                compte = str(compte_raw or "").strip()
                if not compte or not compte[0].isdigit():
                    continue

                # Keep only rows that have at least some descriptive label to avoid structural noise.
                label_raw = ws.cell(row=row_idx, column=label_col).value
                label_text = str(label_raw or "").strip()
                if not label_text:
                    continue

                opening_debit = self._first_numeric_cell(ws, row_idx, opening_debit_cols)
                opening_credit = self._first_numeric_cell(ws, row_idx, opening_credit_cols)
                movement_debit = self._first_numeric_cell(ws, row_idx, movement_debit_cols)
                movement_credit = self._first_numeric_cell(ws, row_idx, movement_credit_cols)
                closing_debit = self._first_numeric_cell(ws, row_idx, closing_debit_cols)
                closing_credit = self._first_numeric_cell(ws, row_idx, closing_credit_cols)

                opening_signed = opening_debit - opening_credit
                closing_signed = closing_debit - closing_credit
                flux_val = self._first_numeric_cell(ws, row_idx, self._resolve_override_columns("flux_columns", []))

                slices[compte] = {
                    "opening_debit": opening_debit,
                    "opening_credit": opening_credit,
                    "movement_debit": movement_debit,
                    "movement_credit": movement_credit,
                    "closing_debit": closing_debit,
                    "closing_credit": closing_credit,
                    "opening_signed": opening_signed,
                    "closing_signed": closing_signed,
                    "flux": flux_val,
                }
            return slices
        finally:
            wb.close()

    def _resolve_override_columns(self, key: str, default: List[int]) -> List[int]:
        value = self.column_overrides.get(key, default)
        if value is None:
            return []
        if isinstance(value, int):
            return [value]
        if isinstance(value, (list, tuple, set)):
            return [int(v) for v in value if v]
        try:
            return [int(value)]
        except (TypeError, ValueError):
            return list(default)

    def _first_numeric_cell(self, ws, row_idx: int, columns: List[int]) -> Decimal:
        if not columns:
            return Decimal("0")
        for col_idx in columns:
            try:
                value = ws.cell(row=row_idx, column=col_idx).value
            except Exception:
                continue
            dec = self._as_decimal(value)
            if dec != 0:
                return dec
        return Decimal("0")

    def _infer_year_hint(self, col_label: Optional[str]) -> Optional[str]:
        if not col_label:
            return None
        label = self._normalize_text(col_label)
        compact = label.replace(" ", "")
        if "n-1" in col_label.lower() or "n - 1" in col_label.lower() or "n1" in compact or "n1" in label:
            return "n1"
        if "n-2" in col_label.lower() or "n - 2" in col_label.lower() or "n2" in compact:
            return "n2"
        if "exercice" in label or "n " in label or label.endswith(" n"):
            return "n"
        return None

    def _infer_value_source(self, col_label: Optional[str]) -> str:
        if not col_label:
            return "final"
        label = self._normalize_text(col_label)
        if any(k in label for k in ("acquisition", "apport", "creation", "augmentation")):
            return "movement_debit"
        if any(k in label for k in ("cession", "sortie", "diminution", "reduction", "scission", "hors service")):
            return "movement_credit"
        if "variation" in label or "ecart" in label:
            return "variation"
        if "flux" in label:
            return "flux"
        if any(k in label for k in ("ouverture", "initial", "n-1")):
            return "opening"
        return "final"

    def _expected_classes_for_sheet(self, sheet_name: str, cell_ref: str) -> Optional[Set[str]]:
        upper = sheet_name.upper()
        col_letter = re.match(r"[A-Z]+", cell_ref).group(0) if re.match(r"[A-Z]+", cell_ref) else ""
        if upper == "BILAN PAYSAGE":
            if col_letter in {"D", "E", "F", "G"}:
                return {"2", "3", "4", "5"}
            if col_letter in {"K", "L"}:
                return {"1", "4", "5"}
        if upper == "COMPTE DE RESULTAT":
            return {"6", "7"}
        if "NOTE 3" in upper:
            return {"2"}
        if "NOTE 4" in upper:
            return {"2", "5"}
        return None

    def _sheet_cell_domain_filter(self, sheet_name: str, cell_ref: str, row: NormalizedBalanceRow) -> bool:
        """Restrict obvious cross-domain leakage by sheet/column."""
        upper = sheet_name.upper()
        col_letter = re.match(r"[A-Z]+", cell_ref).group(0) if re.match(r"[A-Z]+", cell_ref) else ""
        if upper == "BILAN PAYSAGE":
            if col_letter in {"D", "E", "F", "G"} and row.classe not in {"2", "3", "4", "5"}:
                return False
            if col_letter in {"K", "L"} and row.classe not in {"1", "4", "5"}:
                return False
        if upper == "COMPTE DE RESULTAT" and row.classe not in {"6", "7"}:
            return False
        return True

    def _sheet_group(self, sheet_name: str) -> str:
        upper = sheet_name.upper().strip()
        if upper == "BILAN PAYSAGE":
            return "BILAN"
        if upper == "COMPTE DE RESULTAT":
            return "CR"
        if upper == "TABLEAU DES FLUX DE TRESORERIE":
            return "FLUX"
        if upper == "NOTE 13":
            return "NOTE13"
        if upper.startswith("NOTE 3") or upper.startswith("NOTE  3"):
            return "NOTE3"
        if upper.startswith("NOTE 4"):
            return "NOTE4"
        return "OTHER"

    def _build_business_rules(self) -> List[BusinessRule]:
        return [
            # BILAN + NOTES IMMOS (classes 2/3/4/5)
            BusinessRule(("BILAN", "NOTE3"), ("frais de developpement",), ("201",), ("2",), False, 0.62),
            BusinessRule(("BILAN", "NOTE3"), ("brevet", "logiciel"), ("202",), ("2",), False, 0.62),
            BusinessRule(("BILAN", "NOTE3"), ("fonds commercial",), ("203",), ("2",), False, 0.62),
            BusinessRule(("BILAN", "NOTE3"), ("autres immobilisations incorporelles",), ("204",), ("2",), False, 0.60),
            BusinessRule(("BILAN", "NOTE3"), ("immobilisations incorporelles",), ("20",), ("2",), True, 0.56, 18),
            BusinessRule(("BILAN", "NOTE3"), ("terrains",), ("21", "22"), ("2",), True, 0.56, 10),
            BusinessRule(("BILAN", "NOTE3"), ("batiments",), ("22",), ("2",), True, 0.56, 10),
            BusinessRule(("BILAN", "NOTE3"), ("amenagements", "agencements", "installations"), ("23",), ("2",), True, 0.56, 12),
            BusinessRule(("BILAN", "NOTE3"), ("materiel de transport",), ("245", "24"), ("2",), True, 0.56, 12),
            BusinessRule(("BILAN", "NOTE3"), ("materiel", "mobilier"), ("24", "25"), ("2",), True, 0.56, 16),
            BusinessRule(("BILAN", "NOTE3"), ("avances", "acomptes"), ("23",), ("2",), True, 0.56, 10),
            BusinessRule(("BILAN", "NOTE3", "NOTE4"), ("titres de participation",), ("26",), ("2",), True, 0.56, 12),
            BusinessRule(("BILAN", "NOTE4"), ("autres immobilisations financieres",), ("27",), ("2",), True, 0.56, 12),
            BusinessRule(("BILAN", "NOTE4"), ("prets et creances",), ("27", "41", "42", "44", "46"), ("2", "4"), True, 0.54, 16),
            BusinessRule(("BILAN",), ("stocks", "encours"), ("3",), ("3",), True, 0.56, 20),
            BusinessRule(("BILAN",), ("clients",), ("41",), ("4",), True, 0.56, 16),
            BusinessRule(("BILAN",), ("creances", "emplois assimiles"), ("41", "42", "43", "44", "46", "47"), ("4",), True, 0.54, 24),
            BusinessRule(("BILAN",), ("autres creances",), ("46", "47", "48"), ("4",), True, 0.54, 16),
            BusinessRule(("BILAN",), ("titres de placement",), ("50",), ("5",), True, 0.56, 12),
            BusinessRule(("BILAN",), ("valeurs a encaisser",), ("51",), ("5",), True, 0.56, 12),
            BusinessRule(("BILAN",), ("banques", "caisse"), ("52", "53", "54", "56", "57"), ("5",), True, 0.54, 20),
            BusinessRule(("BILAN",), ("total actif immobilise",), ("20", "21", "22", "23", "24", "25", "26", "27"), ("2",), True, 0.52, 24),
            BusinessRule(("BILAN",), ("total actif circulant",), ("3", "4"), ("3", "4"), True, 0.52, 26),
            BusinessRule(("BILAN",), ("total tresorerie actif",), ("5",), ("5",), True, 0.52, 20),
            BusinessRule(("BILAN",), ("total general",), ("2", "3", "4", "5"), ("2", "3", "4", "5"), True, 0.50, 30),
            # COMPTE DE RESULTAT (classes 6/7)
            BusinessRule(("CR",), ("vente de marchandises",), ("70",), ("7",), True, 0.58, 16),
            BusinessRule(("CR",), ("vente de produits fabriques",), ("70", "71"), ("7",), True, 0.58, 16),
            BusinessRule(("CR",), ("travaux", "services vendus"), ("70", "71"), ("7",), True, 0.58, 16),
            BusinessRule(("CR",), ("produits accessoires",), ("70", "71", "72"), ("7",), True, 0.56, 16),
            BusinessRule(("CR",), ("chiffre d affaires",), ("70", "71", "72"), ("7",), True, 0.54, 24),
            BusinessRule(("CR",), ("achat de marchandises",), ("60",), ("6",), True, 0.58, 16),
            BusinessRule(("CR",), ("achats de matieres premieres",), ("60", "61"), ("6",), True, 0.56, 18),
            BusinessRule(("CR",), ("autres achats",), ("60", "61", "62"), ("6",), True, 0.56, 18),
            BusinessRule(("CR",), ("services exterieurs",), ("62",), ("6",), True, 0.58, 14),
            BusinessRule(("CR",), ("transports",), ("62",), ("6",), True, 0.58, 14),
            BusinessRule(("CR",), ("impots et taxes",), ("63",), ("6",), True, 0.58, 14),
            BusinessRule(("CR",), ("charges de personnel",), ("64",), ("6",), True, 0.58, 14),
            BusinessRule(("CR",), ("dotations aux amortissements",), ("68", "69"), ("6",), True, 0.56, 18),
            BusinessRule(("CR",), ("revenus financiers",), ("75", "76", "77"), ("7",), True, 0.56, 18),
            BusinessRule(("CR",), ("frais financiers",), ("67",), ("6",), True, 0.56, 14),
            BusinessRule(("CR",), ("autres produits",), ("75", "76", "77"), ("7",), True, 0.56, 18),
            BusinessRule(("CR",), ("autres charges",), ("65", "66", "67", "68"), ("6",), True, 0.54, 20),
            # TABLEAU DES FLUX (lignes opérationnelles uniquement)
            BusinessRule(("FLUX",), ("capacite d autofinancement globale",), ("6", "7"), ("6", "7"), True, 0.56, 26, True),
            BusinessRule(("FLUX",), ("actif circulant hao",), ("4",), ("4",), True, 0.54, 18, False, "variation"),
            BusinessRule(("FLUX",), ("variation des stocks",), ("3",), ("3",), True, 0.56, 18, False, "variation"),
            BusinessRule(("FLUX",), ("variation des creances",), ("41", "42", "43", "44", "46", "47"), ("4",), True, 0.54, 22, False, "variation"),
            BusinessRule(("FLUX",), ("variation du passif circulant",), ("40", "42", "43", "44", "47"), ("4",), True, 0.54, 22, False, "variation"),
            BusinessRule(("FLUX",), ("variation du bf",), ("3", "4"), ("3", "4"), True, 0.52, 26, False, "variation"),
            BusinessRule(("FLUX",), ("decaissements", "acquisitions", "immobilisation incorporelles"), ("20",), ("2",), True, 0.56, 14, False, "movement_debit"),
            BusinessRule(("FLUX",), ("decaissements", "acquisitions", "immobilisation corporelles"), ("21", "22", "23", "24", "25"), ("2",), True, 0.56, 20, False, "movement_debit"),
            BusinessRule(("FLUX",), ("decaissements", "acquisitions", "immobilisation financieres"), ("26", "27"), ("2",), True, 0.56, 16, False, "movement_debit"),
            BusinessRule(("FLUX",), ("encaissement", "cessions", "immobilisations incorporelles"), ("20", "21", "22", "23", "24", "25"), ("2",), True, 0.56, 20, False, "movement_credit"),
            BusinessRule(("FLUX",), ("encaissement", "cessions", "immobilisations financieres"), ("26", "27"), ("2",), True, 0.56, 16, False, "movement_credit"),
            BusinessRule(("FLUX",), ("augmentation de capital",), ("10",), ("1",), True, 0.56, 12, False, "movement_credit", True),
            BusinessRule(("FLUX",), ("subventions d investissement recues",), ("15",), ("1",), True, 0.56, 12, False, "movement_credit"),
            BusinessRule(("FLUX",), ("prelevement sur le capital",), ("10",), ("1",), True, 0.56, 12, False, "movement_debit", True),
            BusinessRule(("FLUX",), ("dividendes verses",), ("13",), ("1",), True, 0.56, 12, False, "movement_debit"),
            BusinessRule(("FLUX",), ("tresorerie provenant du financement par les capitaux etrangers",), ("16", "17", "47"), ("1", "4"), True, 0.52, 18, False, "final"),
            BusinessRule(("FLUX",), ("emprunts",), ("16", "17"), ("1",), True, 0.56, 14, False, "movement_credit", True),
            BusinessRule(("FLUX",), ("autres dettes financieres",), ("16", "17", "47"), ("1", "4"), True, 0.54, 18, False, "movement_credit", True),
            BusinessRule(("FLUX",), ("remboursemement des emprunts",), ("16", "17"), ("1",), True, 0.56, 14, False, "movement_debit", True),
            # NOTE 13 (capital non appelé / cessions-remboursements)
            BusinessRule(("NOTE13",), ("apporteurs capital non appele",), ("1013", "101300"), ("1",), True, 0.56, 12, False, "movement_debit", True),
            BusinessRule(("NOTE13",), ("total",), ("1013", "101300"), ("1",), True, 0.52, 14, False, "movement_debit", True),
        ]

    def _find_business_rule(self, sheet_name: str, row_label: str) -> Optional[BusinessRule]:
        group = self._sheet_group(sheet_name)
        row_norm = self._normalize_text(row_label)
        if len(row_norm) < 4:
            return None
        for rule in self.business_rules:
            if group not in rule.sheet_groups:
                continue
            if all(keyword in row_norm for keyword in rule.label_keywords):
                return rule
        return None

    def _match_with_business_rules(
        self,
        row_label: str,
        sheet_name: str,
        cell_ref: str,
        col_type: Optional[str],
        year_hint: Optional[str],
        value_source: str,
        expected_classes: Optional[Set[str]],
    ) -> List[MatchedAccount]:
        rule = self._find_business_rule(sheet_name, row_label)
        if not rule:
            return []

        prefixes = set(rule.prefixes)
        classes = set(rule.classes)
        row_tokens = self._tokenize(row_label)
        row_norm = self._normalize_text(row_label)
        effective_classes: Optional[Set[str]] = None
        if expected_classes and classes:
            effective_classes = set(expected_classes) & classes
        elif expected_classes:
            effective_classes = set(expected_classes)
        elif classes:
            effective_classes = classes
        effective_value_source = rule.value_source_override or value_source
        current_group = self._sheet_group(sheet_name)

        candidates: List[Tuple[float, MatchedAccount]] = []
        for compte, normalized_row in self.balance_accounts.items():
            if effective_classes and normalized_row.classe not in effective_classes:
                continue
            if prefixes and not any(compte.startswith(prefix) for prefix in prefixes):
                continue
            if not self._sheet_cell_domain_filter(sheet_name, cell_ref, normalized_row):
                continue

            amount = self._resolve_amount(normalized_row, col_type, year_hint, effective_value_source)
            if amount <= 0 and rule.allow_stock_fallback_when_zero:
                fallback_source = "opening" if year_hint == "n1" else "final"
                amount = self._resolve_amount(normalized_row, col_type, year_hint, fallback_source)
            if amount <= 0 and year_hint == "n1" and rule.allow_n1_fallback:
                amount = self._resolve_amount(normalized_row, col_type, None, "final")
            if amount <= 0:
                continue

            score = self._score_account_match(row_label, row_tokens, normalized_row, row_norm=row_norm)
            if prefixes and any(compte.startswith(prefix) for prefix in prefixes):
                score += 0.08
            score = min(1.0, score)
            min_required = max(0.50, rule.min_score - 0.10) if rule.aggregate else rule.min_score
            # For deterministic aggregate rules with strict SYSCOHADA prefixes (typical FLUX/NOTE13 lines),
            # lexical similarity can be low while mapping is still valid.
            if rule.aggregate and prefixes and current_group in {"FLUX", "NOTE13"}:
                min_required = min(min_required, 0.0)
            if score < min_required:
                continue

            if current_group in {"FLUX", "NOTE13"} and rule.aggregate and prefixes:
                # Confidence floor for deterministic prefix-based rules.
                score = max(score, 0.55)

            candidates.append(
                (
                    score,
                    MatchedAccount(
                        compte=compte,
                        label=normalized_row.label,
                        amount=amount,
                        side=col_type or normalized_row.natural_side,
                        similarity_score=score,
                    ),
                )
            )

        if not candidates:
            return []

        if rule.aggregate:
            candidates.sort(key=lambda item: (-float(item[1].amount), -item[0], item[1].compte))
            selected = [m for _, m in candidates[: rule.max_accounts]]
            return selected

        candidates.sort(key=lambda item: (-item[0], item[1].compte))
        if len(candidates) > 1:
            gap = candidates[0][0] - candidates[1][0]
            if gap < 0.05 and candidates[0][0] < 0.90:
                return []
        return [candidates[0][1]]

    def _preferred_prefixes_for_label(self, row_label: str) -> Optional[Set[str]]:
        label = self._normalize_text(row_label)
        prefix_map = {
            "immobilisations incorporelles": {"20"},
            "immobilisations corporelles": {"21", "22", "23"},
            "titres de participation": {"24"},
            "stocks": {"3"},
            "creances": {"41"},
            "fournisseurs": {"40"},
            "personnel": {"42"},
            "impots": {"44"},
            "banques": {"52"},
            "caisse": {"53"},
            "chiffre d affaires": {"70"},
            "ventes": {"70"},
            "achat": {"60"},
        }
        for key, prefixes in prefix_map.items():
            if key in label:
                return set(prefixes)
        return None

    def _looks_like_section_heading(self, row_label: str) -> bool:
        text = row_label.strip()
        if len(text) < 8:
            return False
        alpha = [c for c in text if c.isalpha()]
        if not alpha:
            return False
        upper_ratio = sum(1 for c in alpha if c.isupper()) / len(alpha)
        return upper_ratio > 0.85

    def _normalize_text(self, value: str) -> str:
        text = unicodedata.normalize("NFKD", value or "")
        text = "".join(ch for ch in text if not unicodedata.combining(ch))
        text = text.lower()
        text = re.sub(r"[^a-z0-9]+", " ", text)
        return " ".join(text.split())

    def _tokenize(self, value: str) -> Set[str]:
        stop_words = {
            "de", "la", "le", "les", "des", "du", "et", "au", "aux", "a", "an", "the",
            "pour", "sur", "par", "en", "d", "l",
        }
        tokens = {tok for tok in self._normalize_text(value).split() if len(tok) > 2 and tok not in stop_words}
        return tokens

    def _token_overlap(self, left: Set[str], right: Set[str]) -> float:
        if not left or not right:
            return 0.0
        inter = len(left & right)
        union = len(left | right)
        return inter / union if union else 0.0

    def _is_structural_text(self, text: str) -> bool:
        if not text:
            return True
        if text in {"ref", "note", "libelles", "libelle", "rubriques", "rubrique"}:
            return True
        if re.fullmatch(r"[a-z0-9]{1,4}", text):
            return True
        return False

    def _as_decimal(self, value: object) -> Decimal:
        if value is None:
            return Decimal("0")
        if isinstance(value, Decimal):
            return value
        if isinstance(value, (int, float)):
            return Decimal(str(value))
        try:
            txt = str(value).strip().replace(" ", "")
            txt = txt.replace(",", ".") if txt.count(",") == 1 and "." not in txt else txt
            return Decimal(txt or "0")
        except (InvalidOperation, ValueError, TypeError):
            return Decimal("0")

    def _keyword_match(self, label1: str, label2: str) -> bool:
        """Check if labels share significant keywords"""
        # Extract keywords (words > 3 chars, skip stop words)
        stop_words = {'de', 'la', 'le', 'des', 'du', 'et', 'or', 'the', 'a', 'an'}
        words1 = set(w for w in label1.split() if len(w) > 3 and w not in stop_words)
        words2 =set(w for w in label2.split() if len(w) > 3 and w not in stop_words)
        
        # If any keyword overlaps, it's a match
        return len(words1 & words2) > 0

    def _hierarchical_match(self, row_label: str) -> Optional[Tuple[str, float]]:
        """
        HYBRID STRATEGY 1: Hierarchical matching
        If 'STOCKS' doesn't fuzzy-match exactly, find accounts that START WITH 'STOCKS'
        """
        if not self.enable_hybrid:
            return None
        
        label_words = row_label.split()
        if not label_words:
            return None
        
        prefix = label_words[0].upper()
        best_match = None
        best_score = 0.35  # Lower threshold for hierarchical matching
        
        for compte, normalized_row in self.balance_accounts.items():
            if normalized_row.label.upper().startswith(prefix):
                # Calculate partial similarity
                ratio = SequenceMatcher(None, row_label.lower(), normalized_row.label.lower()).ratio()
                if ratio > best_score:
                    best_match = (compte, ratio)
                    best_score = ratio
        
        return best_match
    
    def _pattern_match(self, row_label: str) -> Optional[Tuple[str, float]]:
        """
        HYBRID STRATEGY 2: Pattern matching
        Try to find accounts that contain any significant word from row_label
        """
        if not self.enable_hybrid:
            return None
        
        label_upper = row_label.upper()
        words = [w for w in row_label.split() if len(w) > 4]  # Significant words
        
        if not words:
            return None
        
        best_match = None
        best_score = 0.45
        
        for compte, normalized_row in self.balance_accounts.items():
            account_upper = normalized_row.label.upper()
            
            # Check if any word from label is in account label
            for word in words:
                if word in account_upper:
                    ratio = 0.60  # Moderate confidence for word match
                    if ratio > best_score:
                        best_match = (compte, ratio)
                        best_score = ratio
                    break
        
        return best_match

    def _get_fallback_matches(self, row_label: str, col_type: Optional[str]) -> List[MatchedAccount]:
        """
        Collect matches using hybrid strategies (hierarchical, pattern)
        Only called if fuzzy matching returns no results
        """
        matches: List[Tuple[float, MatchedAccount]] = []
        
        # Try hierarchical matching
        hier_result = self._hierarchical_match(row_label)
        if hier_result:
            compte, similarity = hier_result
            normalized_row = self.balance_accounts[compte]
            
            # Get appropriate amount based on column type
            amount = Decimal("0")
            if col_type == "debit":
                amount = normalized_row.debit_balance
            elif col_type == "credit":
                amount = normalized_row.credit_balance
            else:
                if normalized_row.natural_side == "debit":
                    amount = normalized_row.debit_balance
                else:
                    amount = normalized_row.credit_balance
            
            if amount > 0:
                matched = MatchedAccount(
                    compte=compte,
                    label=normalized_row.label,
                    amount=amount,
                    side=col_type or normalized_row.natural_side,
                    similarity_score=similarity,
                )
                matches.append((similarity, matched))
                logger.info(f"Hierarchical match for '{row_label}': {compte} (score={similarity:.2f})")
        
        # Try pattern matching
        if not matches:
            pattern_result = self._pattern_match(row_label)
            if pattern_result:
                compte, similarity = pattern_result
                normalized_row = self.balance_accounts[compte]
                
                # Get appropriate amount
                amount = Decimal("0")
                if col_type == "debit":
                    amount = normalized_row.debit_balance
                elif col_type == "credit":
                    amount = normalized_row.credit_balance
                else:
                    if normalized_row.natural_side == "debit":
                        amount = normalized_row.debit_balance
                    else:
                        amount = normalized_row.credit_balance
                
                if amount > 0:
                    matched = MatchedAccount(
                        compte=compte,
                        label=normalized_row.label,
                        amount=amount,
                        side=col_type or normalized_row.natural_side,
                        similarity_score=similarity,
                    )
                    matches.append((similarity, matched))
                    logger.info(f"Pattern match for '{row_label}': {compte} (score={similarity:.2f})")
        
        return [m[1] for m in sorted(matches, key=lambda x: -x[0])]

    def _fuzzy_similarity(self, str1: str, str2: str) -> float:
        """
        Calculate fuzzy similarity between two strings.
        Returns value 0.0-1.0
        """
        return SequenceMatcher(None, str1, str2).ratio()

    def _extract_column_type(self, col_label: Optional[str]) -> Optional[str]:
        """Extract 'debit' or 'credit' from column label (French & English)"""
        if not col_label:
            return None

        col_lower = col_label.lower().strip()
        
        # French variations
        if any(word in col_lower for word in ["débit", "debit", "db", "d:"]):
            return "debit"
        if any(word in col_lower for word in ["crédit", "credit", "cr", "c:"]):
            return "credit"
        
        # English variations
        if "debit" in col_lower or "dr" in col_lower:
            return "debit"
        if "credit" in col_lower or "cr" in col_lower:
            return "credit"
            
        return None

    def _write_cell(self, ws, need: CellNeed, assignment: CellAssignment) -> bool:
        """Write assignment value to cell (handle merged cells safely)"""
        try:
            # Find the actual cell to write to
            actual_cell = ws[need.cell]

            # Check if it's part of a merged range and use master cell
            for merged_range in ws.merged_cells.ranges:
                if need.cell in merged_range:
                    # Use the master cell (top-left of merged range)
                    actual_cell = ws[merged_range.start_cell.coordinate]
                    break

            # Write the value
            actual_cell.value = float(assignment.total_amount)
            return True
            
        except Exception as e:
            logger.debug(f"Failed to write {need.cell}: {e}")
            return False
    
    def _apply_column_formulas(self) -> None:
        """
        Apply formulas to columns that have formula definitions in headers.
        This is called after filling values, to insert Excel formulas in computed columns.
        """
        for sheet_name, detector in self.formula_detectors.items():
            ws = self.wb[sheet_name]
            formula_columns = detector.get_formula_columns()
            formula_seeds = self._collect_formula_seeds(ws)

            if not formula_columns and not formula_seeds:
                continue

            if formula_columns:
                logger.info(f"  Sheet {sheet_name}: applying formulas to columns {', '.join(formula_columns)}")

            # Determine data range (skip header rows, process data rows)
            data_start_row = detector.header_row_detected + 1 if detector.header_row_detected else 11

            # Process each data row
            for row_idx in range(data_start_row, ws.max_row + 1):
                # Check if this is a data row (has content in first column)
                first_cell = ws.cell(row=row_idx, column=1)
                if not first_cell.value:
                    continue

                # Skip if it's a header or total row
                if isinstance(first_cell.value, str):
                    text_lower = first_cell.value.lower()
                    if any(kw in text_lower for kw in ['total', 'sous-total', 'en-tête', 'header', 'libellé']):
                        continue

                # Apply formulas detected from headers
                for col_letter in formula_columns:
                    cell = ws[f"{col_letter}{row_idx}"]
                    if cell.data_type == "f":
                        continue
                    if cell.value is None or isinstance(cell.value, (int, float)):
                        formula = detector.apply_formula_to_cell(col_letter, row_idx)
                        if formula:
                            try:
                                cell.value = formula
                                self.formulas_applied += 1
                                self.formula_logs.append(
                                    {
                                        "sheet": sheet_name,
                                        "cell": cell.coordinate,
                                        "formula": str(formula),
                                        "source": "header",
                                    }
                                )
                                logger.debug(f"    {cell.coordinate}: {formula}")
                            except Exception as e:
                                logger.warning(f"Failed to apply formula to {cell.coordinate}: {e}")

                # Apply formulas from template seeds (if any)
                for col_letter, (seed_cell, seed_formula) in formula_seeds.items():
                    target_cell = ws[f"{col_letter}{row_idx}"]
                    if target_cell.data_type == "f":
                        continue
                    if target_cell.value is None or isinstance(target_cell.value, (int, float)):
                        try:
                            translated = Translator(seed_formula, origin=seed_cell).translate_formula(target_cell.coordinate)
                            target_cell.value = translated
                            self.formulas_applied += 1
                            self.formula_logs.append(
                                {
                                    "sheet": sheet_name,
                                    "cell": target_cell.coordinate,
                                    "formula": str(translated),
                                    "source": f"seed:{seed_cell}",
                                }
                            )
                            logger.debug(f"    {target_cell.coordinate}: {translated}")
                        except Exception as e:
                            logger.warning(f"Failed to translate formula to {target_cell.coordinate}: {e}")

    def _collect_formula_seeds(self, ws) -> Dict[str, tuple[str, str]]:
        """Collect one formula seed per column from the template."""
        seeds: Dict[str, tuple[str, str]] = {}
        for row in ws.iter_rows():
            for cell in row:
                if cell.data_type == "f" and isinstance(cell.value, str):
                    col_letter = cell.column_letter
                    if col_letter not in seeds:
                        seeds[col_letter] = (cell.coordinate, cell.value)
        return seeds

    def write_formula_log(self, path: Path | str) -> Path:
        """Write formula application log to JSON file."""
        output = Path(path)
        output.parent.mkdir(parents=True, exist_ok=True)
        payload = {
            "total": len(self.formula_logs),
            "items": self.formula_logs,
        }
        output.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        return output

    def save(self, output_path: Path | str) -> Path:
        """Save filled workbook"""
        if self.wb is None:
            raise RuntimeError("Workbook not loaded")
        output = Path(output_path)
        output.parent.mkdir(parents=True, exist_ok=True)
        try:
            self.wb.save(output)
            logger.info(f"Saved to {output}")
            return output
        except PermissionError as exc:
            from datetime import datetime
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            fallback = output.with_name(f"{output.stem}_{timestamp}{output.suffix}")
            logger.warning(
                "Impossible d'ecrire %s (%s). Sauvegarde de secours vers %s.",
                output,
                exc,
                fallback,
            )
            self.wb.save(fallback)
            logger.info(f"Saved to {fallback}")
            return fallback

    def get_stats(self) -> Dict[str, object]:
        """Get filling statistics"""
        total_assigned = sum(a.total_amount for a in self.assignments)
        return {
            "assignments": len(self.assignments),
            "unmatched_cells": len(self.unmatched_cells),
            "total_amount": float(total_assigned),
            "average_confidence": (
                sum(a.confidence for a in self.assignments) / len(self.assignments)
                if self.assignments
                else 0.0
            ),
            "high_confidence": len([a for a in self.assignments if a.confidence > 0.8]),
        }


__all__ = ["SemanticBalanceFiller", "CellAssignment", "CellNeed"]
