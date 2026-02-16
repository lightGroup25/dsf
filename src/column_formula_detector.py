# -*- coding: utf-8 -*-
"""
Column Formula Detector - Détecte et implémente les formules dans les en-têtes de colonnes
=========================================================================================

Ce module analyse les en-têtes de colonnes pour détecter des formules comme:
- "VARIATION = Clôture - Ouverture"
- "Total = Col1 + Col2"
- "% = (Valeur / Total) * 100"

Et génère automatiquement les formules Excel correspondantes.
"""
from __future__ import annotations

import logging
import re
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import get_column_letter, column_index_from_string

logger = logging.getLogger(__name__)


@dataclass
class ColumnFormula:
    """Représente une formule détectée dans un en-tête de colonne"""
    sheet: str
    column_letter: str
    column_index: int
    header_row: int
    header_text: str
    formula_pattern: str  # Pattern de formule détecté (ex: "{A} - {B}")
    referenced_columns: List[str]  # Colonnes référencées (ex: ["D", "E"])
    operation_type: str  # "subtraction", "addition", "multiplication", "division", "percentage", "complex"
    
    def generate_formula(self, row_number: int) -> str:
        """
        Génère la formule Excel pour une ligne spécifique.
        
        Args:
            row_number: Numéro de ligne pour laquelle générer la formule
            
        Returns:
            Formule Excel (ex: "=E10-D10")
        """
        formula = self.formula_pattern
        
        # Remplacer les références de colonnes par les références avec numéro de ligne
        for col in self.referenced_columns:
            formula = formula.replace(f"{{{col}}}", f"{col}{row_number}")
        
        return f"={formula}"


class ColumnFormulaDetector:
    """
    Détecte les formules dans les en-têtes de colonnes et génère les formules Excel.
    
    Patterns reconnus:
    - "VARIATION = Clôture - Ouverture"
    - "Écart = A - B"
    - "Total = Débit + Crédit"
    - "% = (Valeur / Base) * 100"
    - "Ratio = A / B"
    """
    
    # Patterns de formules reconnus
    FORMULA_PATTERNS = [
        # Pattern: "Label = ColA - ColB" ou "Label = A - B"
        (r'=\s*([A-Za-zÀ-ÿ\s]+)\s*[-−]\s*([A-Za-zÀ-ÿ\s]+)', 'subtraction'),
        # Pattern: "Label = ColA + ColB" ou "Label = A + B"
        (r'=\s*([A-Za-zÀ-ÿ\s]+)\s*\+\s*([A-Za-zÀ-ÿ\s]+)', 'addition'),
        # Pattern: "Label = ColA * ColB"
        (r'=\s*([A-Za-zÀ-ÿ\s]+)\s*[*×]\s*([A-Za-zÀ-ÿ\s]+)', 'multiplication'),
        # Pattern: "Label = ColA / ColB"
        (r'=\s*([A-Za-zÀ-ÿ\s]+)\s*/\s*([A-Za-zÀ-ÿ\s]+)', 'division'),
        # Pattern: "Label = (A / B) * 100" ou "Label = A/B*100"
        (r'=\s*\(?\s*([A-Za-zÀ-ÿ\s]+)\s*/\s*([A-Za-zÀ-ÿ\s]+)\s*\)?\s*\*\s*100', 'percentage'),
    ]
    
    def __init__(self, ws: Worksheet, max_header_row: int = 15):
        """
        Args:
            ws: Feuille Excel à analyser
            max_header_row: Nombre maximum de lignes à scanner pour les en-têtes
        """
        self.ws = ws
        self.sheet_name = ws.title
        self.max_header_row = max_header_row
        self.column_formulas: Dict[str, ColumnFormula] = {}  # {col_letter: ColumnFormula}
        self.column_labels: Dict[str, str] = {}  # {col_letter: label}
        self.header_row_detected: Optional[int] = None
        
    def detect(self) -> Dict[str, ColumnFormula]:
        """
        Détecte toutes les formules dans les en-têtes de colonnes.
        
        Returns:
            Dict mapping column letters to ColumnFormula objects
        """
        logger.info(f"Detecting column formulas in sheet: {self.sheet_name}")
        
        # Étape 1: Trouver la ligne d'en-tête
        self._find_header_row()
        
        if not self.header_row_detected:
            logger.debug(f"No header row detected in {self.sheet_name}")
            return {}
        
        logger.debug(f"Header row detected at line {self.header_row_detected}")
        
        # Étape 2: Extraire les labels de colonnes
        self._extract_column_labels()
        
        # Étape 3: Détecter les formules dans les en-têtes
        self._detect_formulas_in_headers()
        
        logger.info(f"Detected {len(self.column_formulas)} column formulas in {self.sheet_name}")
        
        return self.column_formulas
    
    def _find_header_row(self) -> None:
        """
        Trouve la ligne d'en-tête en cherchant des mots-clés typiques.
        """
        keywords = [
            'ouverture', 'opening', 'clôture', 'closing', 'clture',
            'débit', 'debit', 'crédit', 'credit',
            'variation', 'écart', 'ecart', 'mouvement',
            'libellé', 'libelle', 'label', 'compte', 'account',
            'total', 'montant', 'amount', 'valeur', 'value'
        ]
        
        for row_idx in range(1, min(self.max_header_row + 1, self.ws.max_row + 1)):
            row_text = []
            for col_idx in range(1, min(20, self.ws.max_column + 1)):  # Scan 20 premières colonnes
                cell = self.ws.cell(row=row_idx, column=col_idx)
                if cell.value and isinstance(cell.value, str):
                    row_text.append(cell.value.lower())
            
            # Si au moins 2 mots-clés trouvés dans la ligne
            row_text_combined = ' '.join(row_text)
            keyword_count = sum(1 for kw in keywords if kw in row_text_combined)
            
            if keyword_count >= 2:
                self.header_row_detected = row_idx
                return
    
    def _extract_column_labels(self) -> None:
        """
        Extrait les labels de toutes les colonnes depuis la ligne d'en-tête.
        """
        if not self.header_row_detected:
            return
        
        for col_idx in range(1, self.ws.max_column + 1):
            cell = self.ws.cell(row=self.header_row_detected, column=col_idx)
            if cell.value and isinstance(cell.value, str):
                col_letter = get_column_letter(col_idx)
                self.column_labels[col_letter] = cell.value.strip()
    
    def _detect_formulas_in_headers(self) -> None:
        """
        Détecte les formules dans les en-têtes de colonnes.
        """
        for col_letter, header_text in self.column_labels.items():
            formula = self._parse_formula_from_text(col_letter, header_text)
            if formula:
                self.column_formulas[col_letter] = formula
                logger.debug(f"Formula detected in column {col_letter}: {header_text}")
    
    def _parse_formula_from_text(self, col_letter: str, text: str) -> Optional[ColumnFormula]:
        """
        Parse un texte d'en-tête pour extraire une formule.
        
        Args:
            col_letter: Lettre de la colonne
            text: Texte de l'en-tête
            
        Returns:
            ColumnFormula si une formule est détectée, None sinon
        """
        text_lower = text.lower()
        
        # Essayer chaque pattern
        for pattern, operation_type in self.FORMULA_PATTERNS:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                # Extraire les colonnes référencées
                ref1 = match.group(1).strip()
                ref2 = match.group(2).strip()
                
                # Convertir les noms de colonnes en lettres
                col1 = self._resolve_column_reference(ref1)
                col2 = self._resolve_column_reference(ref2)
                
                if not col1 or not col2:
                    continue
                
                # Construire le pattern de formule
                if operation_type == 'subtraction':
                    formula_pattern = f"{{{col1}}}-{{{col2}}}"
                elif operation_type == 'addition':
                    formula_pattern = f"{{{col1}}}+{{{col2}}}"
                elif operation_type == 'multiplication':
                    formula_pattern = f"{{{col1}}}*{{{col2}}}"
                elif operation_type == 'division':
                    formula_pattern = f"{{{col1}}}/{{{col2}}}"
                elif operation_type == 'percentage':
                    formula_pattern = f"({{{col1}}}/{{{col2}}})*100"
                else:
                    continue
                
                return ColumnFormula(
                    sheet=self.sheet_name,
                    column_letter=col_letter,
                    column_index=column_index_from_string(col_letter),
                    header_row=self.header_row_detected,
                    header_text=text,
                    formula_pattern=formula_pattern,
                    referenced_columns=[col1, col2],
                    operation_type=operation_type
                )
        
        return None
    
    def _resolve_column_reference(self, ref_text: str) -> Optional[str]:
        """
        Résout une référence de colonne depuis un texte.
        
        Essaie plusieurs stratégies:
        1. Si c'est déjà une lettre de colonne (A, B, AA, etc.) → retourner tel quel
        2. Si c'est un label connu → trouver la colonne correspondante
        3. Fuzzy match avec les labels existants
        
        Args:
            ref_text: Texte de référence (ex: "Clôture", "A", "Solde")
            
        Returns:
            Lettre de colonne (ex: "D") ou None
        """
        ref_clean = ref_text.strip().upper()
        
        # Stratégie 1: C'est déjà une lettre de colonne
        if re.match(r'^[A-Z]{1,3}$', ref_clean):
            return ref_clean
        
        # Stratégie 2: Match exact avec un label
        ref_lower = ref_text.strip().lower()
        for col_letter, label in self.column_labels.items():
            if ref_lower in label.lower() or label.lower() in ref_lower:
                return col_letter
        
        # Stratégie 3: Fuzzy match
        from difflib import SequenceMatcher
        best_match = None
        best_score = 0.0
        
        for col_letter, label in self.column_labels.items():
            similarity = SequenceMatcher(None, ref_lower, label.lower()).ratio()
            if similarity > best_score and similarity > 0.6:  # Seuil 60%
                best_score = similarity
                best_match = col_letter
        
        return best_match
    
    def apply_formula_to_cell(self, col_letter: str, row_number: int) -> Optional[str]:
        """
        Génère la formule Excel pour une cellule spécifique.
        
        Args:
            col_letter: Lettre de la colonne
            row_number: Numéro de ligne
            
        Returns:
            Formule Excel (ex: "=E10-D10") ou None si pas de formule
        """
        if col_letter not in self.column_formulas:
            return None
        
        formula = self.column_formulas[col_letter]
        return formula.generate_formula(row_number)
    
    def get_formula_columns(self) -> List[str]:
        """Retourne la liste des lettres de colonnes qui ont des formules."""
        return list(self.column_formulas.keys())
    
    def has_formula(self, col_letter: str) -> bool:
        """Vérifie si une colonne a une formule."""
        return col_letter in self.column_formulas


__all__ = ["ColumnFormulaDetector", "ColumnFormula"]
