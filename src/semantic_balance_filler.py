# -*- coding: utf-8 -*-
"""
Semantic Balance Filler - Intelligent cell-by-cell matching
Analyzes cell labels (row + column), fuzzy-matches balance accounts, and fills values.
"""
from __future__ import annotations

import json
import logging
from dataclasses import dataclass, field
from decimal import Decimal
from difflib import SequenceMatcher
from pathlib import Path
from typing import Dict, List, Optional, Tuple

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


class SemanticBalanceFiller:
    def __init__(
        self,
        template_path: Path | str,
        balance_file: Path | str,
        inventory: DSFInventory,
        fuzzy_threshold: float = 0.70,
        enable_hybrid: bool = False,
    ):
        self.template_path = Path(template_path)
        self.balance_file = Path(balance_file)
        self.inventory = inventory
        self.fuzzy_threshold = fuzzy_threshold  # Augmenté de 0.6 à 0.70 pour plus de stricte
        self.enable_hybrid = enable_hybrid  # NEW: Enable hybrid optimization strategies
        self.wb = None
        self.balance_accounts: Dict[str, NormalizedBalanceRow] = {}
        self.assignments: List[CellAssignment] = []
        self.unmatched_cells: List[CellNeed] = []
        self.formula_detectors: Dict[str, ColumnFormulaDetector] = {}  # {sheet_name: detector}
        self.formulas_applied: int = 0
        self.formula_logs: List[Dict[str, str]] = []

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
        
        # Workbook already loaded above.
        
        # Manual column mapping for this specific balance file structure
        # Row structure:
        # Row 8: Section headers (Mouvements au 31/12/23, Mouvements, Soldes cumulés)
        # Row 10: Debit/Credit labels
        # Data starts at row 13
        # VERIFIED: Real data is in columns 14 (debit) and 27 (credit), not 23/26
        column_overrides = {
            'compte': 1,           # C1: Numéro de compte (row 8)
            'label': 4,            # C4: Intitulé des comptes (row 9)
            'debit': 14,           # C14: Soldes - Débit (VERIFIED with data at row 25)
            'credit': 27,          # C27: Soldes - Crédit (VERIFIED with data at row 25)
        }
        
        # Load balance accounts with explicit column mapping
        normalizer = BalanceNormalizer(self.balance_file, column_overrides=column_overrides)
        for normalized_row in normalizer.iterate():
            self.balance_accounts[normalized_row.compte] = normalized_row
        logger.info(f"Loaded {len(self.balance_accounts)} accounts from balance")
        
        # Debug: show first few accounts with amounts
        if self.balance_accounts:
            for i, (compte, row) in enumerate(list(self.balance_accounts.items())[:3]):
                logger.info(f"  [{i}] {compte}: debit={row.debit_balance}, credit={row.credit_balance}, natural_side={row.natural_side}")
        
        # Count accounts with non-zero amounts
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
        
        # Feuilles à exclure (informations générales, pas de données de balance)
        EXCLUDED_SHEETS = ["ENTETE", "ENTÊTE", "Fiche R1", "R1", "Fiche R2", "R2", "Fiche R3", "R3"]

        for sheet_name in self.wb.sheetnames:
            # SKIP fiches d'informations générales
            if sheet_name in EXCLUDED_SHEETS:
                logger.info(f"Skipping general info sheet: {sheet_name}")
                continue
            
            ws = self.wb[sheet_name]
            print(f"  Processing sheet: {sheet_name}")  # Debug
            logger.info(f"Processing sheet: {sheet_name}")
            
            # Check if this sheet has formula columns
            has_formula_detector = sheet_name in self.formula_detectors
            formula_detector = self.formula_detectors.get(sheet_name) if has_formula_detector else None
            
            # Build merged cells set for fast lookup
            merged_coords = set()
            for merged_range in ws.merged_cells.ranges:
                for cell in merged_range.cells:
                    merged_coords.add(f"{cell[0]}{cell[1]}")
            
            print(f"    Merged cells: {len(merged_coords)}, Max row: {ws.max_row}")  # Debug

            # Iterate rows (skip first 10 typically used for headers/structure)
            for row_num in range(11, ws.max_row + 1):
                # Get row label (usually column A or B)
                row_label = self._extract_row_label(ws, row_num)
                if not row_label or row_label.strip() == "":
                    continue

                # Iterate columns in this row
                for col_num in range(1, ws.max_column + 1):
                    cell = ws.cell(row=row_num, column=col_num)
                    
                    # FIRST: Skip merged cells completely (they're read-only)
                    if cell.coordinate in merged_coords:
                        continue
                    
                    # Check if this column has a formula definition - if so, skip filling with value
                    col_letter = get_column_letter(col_num)
                    if formula_detector and formula_detector.has_formula(col_letter):
                        logger.debug(f"Skipping {cell.coordinate} - formula column")
                        continue
                    
                    # Inventory-based strict validation (locked/formula/type expected)
                    if not self._is_inventory_writable(sheet_name, cell.coordinate, self._extract_col_label(ws, col_num)):
                        continue

                    # Skip if already has formula/data or is structure cell
                    if not self._is_fillable_cell(cell):
                        continue

                    col_label = self._extract_col_label(ws, col_num)
                    if not col_label:
                        continue

                    # Create need and try to fill
                    need = CellNeed(
                        sheet=sheet_name,
                        cell=cell.coordinate,
                        row_num=row_num,
                        col_num=col_num,
                        row_label=row_label,
                        col_label=col_label,
                        is_merged=False,
                    )

                    assignment = self._satisfy_cell_need(need)
                    if assignment:
                        self.assignments.append(assignment)
                        self._write_cell(ws, need, assignment)
                        assignments_count += 1
                    else:
                        self.unmatched_cells.append(need)

        print(f"  TOTAL FILLED: {assignments_count}")  # Debug
        
        # Now apply column formulas where detected
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

    def _extract_row_label(self, ws, row_num: int) -> Optional[str]:
        """Extract label from first column of row (usually column A or B)"""
        # Try column A first, then B
        for col in [1, 2]:
            cell = ws.cell(row=row_num, column=col)
            if cell.value and isinstance(cell.value, str):
                label = cell.value.strip()
                # Only skip empty strings after strip
                if label:
                    return label
        return None

    def _extract_col_label(self, ws, col_num: int) -> Optional[str]:
        """Extract label from first few rows of column (header rows)"""
        # Look for column label in rows 1-10 (typically headers)
        for row in range(1, 11):
            cell = ws.cell(row=row, column=col_num)
            if cell.value and isinstance(cell.value, str):
                val = cell.value.strip()
                if val and len(val) > 0:
                    return val
        return None

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

    def _satisfy_cell_need(self, need: CellNeed) -> Optional[CellAssignment]:
        """
        Main matching logic:
        1. Create COMPOSITE label from row_label + col_label  
        2. Fuzzy match this composite against balance accounts
        3. Filter by column type (debit/credit)
        4. Accumulate if "Total" in row label and not found in balance
        """
        row_label_lower = need.row_label.lower()
        is_total = "total" in row_label_lower
        col_type = self._extract_column_type(need.col_label)  # "debit" or "credit"

        # CREATE COMPOSITE LABEL from row + col 
        composite_label = f"{need.row_label} {need.col_label}".strip() if need.col_label else need.row_label
        
        # Find matching accounts using BOTH row and column labels
        matches = self._fuzzy_match_accounts(composite_label, col_type, is_total)

        if not matches:
            logger.debug(f"No matches for {need.cell}: row+col composite")
            return None

        # Build assignment
        total_amount = sum(m.amount for m in matches)
        logger.debug(f"Assignment {need.cell}: {len(matches)} account(s), amount={total_amount}")
        
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
        )

        return assignment
        
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
        )

        return assignment

    def _fuzzy_match_accounts(
        self, row_label: str, col_type: Optional[str], is_total: bool
    ) -> List[MatchedAccount]:
        """
        Fuzzy match row label to balance accounts.
        Returns sorted list by similarity score (best first).
        
        STRATEGY CASCADE (if enable_hybrid=True):
        1. Fuzzy matching (threshold >= 0.50)
        2. Hierarchical matching (prefix-based)
        3. Pattern matching (word presence)
        """
        matches: List[Tuple[float, MatchedAccount]] = []

        for compte, normalized_row in self.balance_accounts.items():
            # Calculate fuzzy similarity
            similarity = self._fuzzy_similarity(
                row_label.lower(),
                normalized_row.label.lower()
            )

            # PRIMARY: Try fuzzy match
            if similarity >= self.fuzzy_threshold:
                pass  # Use this match below
            # FALLBACK: Keyword-based match (if row_label contains a key word from account)
            elif self._keyword_match(row_label.lower(), normalized_row.label.lower()):
                similarity = 0.55  # Assign moderate confidence for keyword match
            else:
                continue  # No match

            # Determine which balance to use based on column type OR natural side
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

            # Skip if no amount
            if amount <= 0:
                logger.debug(f"Skipping {compte} ({normalized_row.label}): amount={amount}")
                continue

            matched = MatchedAccount(
                compte=compte,
                label=normalized_row.label,
                amount=amount,
                side=col_type or normalized_row.natural_side,
                similarity_score=similarity,
            )
            matches.append((similarity, matched))
            logger.debug(f"Matched {compte}: similarity={similarity:.2f}, amount={amount}")

        # Sort by similarity (best first)
        matches.sort(key=lambda x: -x[0])

        if matches:
            best_matches = [m[1] for m in matches if m[0] == matches[0][0]]

            # If "total" and only one best match, accumulate similar ones
            if is_total and len(best_matches) == 1:
                return [m[1] for m in matches]
            else:
                return best_matches[:1]
        
        # HYBRID: If no fuzzy match, try hierarchical and pattern matching
        if self.enable_hybrid and not matches:
            logger.debug(f"No fuzzy match for '{row_label}', trying hybrid strategies...")
            fallback_matches = self._get_fallback_matches(row_label, col_type)
            return fallback_matches[:1]
        
        return []

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
