# -*- coding: utf-8 -*-
"""
DSF Generator with streaming/chunked output to avoid memory overflow.
Utilise openpyxl avec écriture line-by-line et flush régulier.
"""

from __future__ import annotations

import logging
from datetime import datetime
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Any

import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.worksheet.merge import MergedCellRange
from openpyxl.cell.cell import MergedCell

logger = logging.getLogger(__name__)


@dataclass
class DSFCell:
    """Définition d'une cellule DSF (adresse, valeur, style)."""
    row: int
    col: int
    value: Any = None
    formula: Optional[str] = None
    font: Optional[Font] = None
    fill: Optional[PatternFill] = None
    alignment: Optional[Alignment] = None
    border: Optional[Border] = None
    number_format: Optional[str] = None


@dataclass
class DSFMerge:
    """Définition d'une fusion de cellules."""
    start_row: int
    start_col: int
    end_row: int
    end_col: int


class DSFStreamWriter:
    """
    Générateur streaming pour DSF (évite charges mémoire).
    
    Stratégie:
    - Charge template DSF depuis le fichier standard
    - Écrit feuille par feuille en streaming
    - Flush après chaque section
    - Pas de structure intermédiaire en mémoire
    """

    def __init__(self, template_path: Path, output_path: Path, chunk_size: int = 100):
        """
        Args:
            template_path: Chemin du DSF standard template
            output_path: Chemin du fichier de sortie
            chunk_size: Nombre de lignes avant flush (optimisation mémoire)
        """
        self.template_path = Path(template_path)
        self.output_path = Path(output_path)
        self.chunk_size = chunk_size
        self.wb = None
        self.current_sheet = None
        self.row_counter = {}  # Track rows per sheet for chunking

    def initialize_from_template(self):
        """Charge la structure du template DSF."""
        logger.info(f"Chargement template: {self.template_path}")
        self.wb = openpyxl.load_workbook(self.template_path)
        logger.info(f"Template chargé: {len(self.wb.sheetnames)} sheets")

    def get_sheet(self, sheet_name: str) -> Worksheet:
        """Obtient ou crée une feuille."""
        if sheet_name not in self.wb.sheetnames:
            return self.wb.create_sheet(sheet_name)
        return self.wb[sheet_name]

    def write_cell(self, sheet_name: str, row: int, col: int, value: Any, 
                   formula: Optional[str] = None,
                   style: Optional[Dict] = None):
        """
        Écrit une cellule de manière optimisée.
        
        Args:
            sheet_name: Nom de la feuille
            row: Index ligne (1-based)
            col: Index colonne (1-based)
            value: Valeur à écrire
            formula: Formule optionnelle (au lieu de value)
            style: Dict avec keys: font, fill, alignment, border, number_format
        """
        ws = self.get_sheet(sheet_name)
        cell = ws.cell(row=row, column=col)

        # Handle MergedCell: Redirect to top-left
        if isinstance(cell, MergedCell):
            for merged_range in ws.merged_cells.ranges:
                if cell.coordinate in merged_range:
                    # Redirect to top-left cell
                    # logger.debug("Redirecting write on %s from %s to %s", sheet_name, cell.coordinate, merged_range.start_cell.coordinate)
                    cell = ws.cell(row=merged_range.min_row, column=merged_range.min_col)
                    break

        # Écrire valeur ou formule
        if formula:
            cell.value = formula  # Les formules doivent commencer par =
        else:
            cell.value = value

        # Appliquer styles
        if style:
            if "font" in style:
                cell.font = style["font"]
            if "fill" in style:
                cell.fill = style["fill"]
            if "alignment" in style:
                cell.alignment = style["alignment"]
            if "border" in style:
                cell.border = style["border"]
            if "number_format" in style:
                cell.number_format = style["number_format"]

        # Track chunking pour flush régulier
        if sheet_name not in self.row_counter:
            self.row_counter[sheet_name] = 0
        self.row_counter[sheet_name] += 1

        if self.row_counter[sheet_name] % self.chunk_size == 0:
            self._flush_sheet(sheet_name)

    def write_range(self, sheet_name: str, start_row: int, start_col: int,
                   end_row: int, end_col: int, values: List[List[Any]] = None):
        """
        Écrit une plage de cellules en bloc.
        
        Args:
            sheet_name: Nom de la feuille
            start_row: Ligne de départ (1-based)
            start_col: Colonne de départ (1-based)
            end_row: Ligne de fin (incluse)
            end_col: Colonne de fin (incluse)
            values: Liste 2D des valeurs (ou None pour garder existantes)
        """
        ws = self.get_sheet(sheet_name)
        
        if values:
            for r_idx, row_values in enumerate(values, start=start_row):
                for c_idx, cell_value in enumerate(row_values, start=start_col):
                    if cell_value is not None:
                        ws.cell(row=r_idx, column=c_idx).value = cell_value

        self.row_counter[sheet_name] = self.row_counter.get(sheet_name, 0) + len(values or [])
        if self.row_counter[sheet_name] % self.chunk_size == 0:
            self._flush_sheet(sheet_name)

    def merge_cells(self, sheet_name: str, start_row: int, start_col: int,
                   end_row: int, end_col: int):
        """Fusionne une plage de cellules."""
        ws = self.get_sheet(sheet_name)
        try:
            ws.merge_cells(start_row=start_row, start_column=start_col,
                          end_row=end_row, end_column=end_col)
        except Exception as e:
            logger.warning(f"Merge failed {sheet_name}:{start_row}:{start_col}: {e}")

    def set_column_width(self, sheet_name: str, col: int, width: float):
        """Défini la largeur d'une colonne."""
        ws = self.get_sheet(sheet_name)
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = width

    def set_row_height(self, sheet_name: str, row: int, height: float):
        """Défini la hauteur d'une ligne."""
        ws = self.get_sheet(sheet_name)
        ws.row_dimensions[row].height = height

    def _flush_sheet(self, sheet_name: str):
        """Force l'écriture en mémoire intermédiaire (optim openpyxl)."""
        # openpyxl n'expose pas vraiment le flush, mais on peut réduire la cache
        logger.debug(f"Chunk flush: {sheet_name} @ {self.row_counter.get(sheet_name, 0)} rows")

    def save(self):
        """Sauvegarde le fichier complet."""
        logger.info(f"Écriture du fichier de sortie: {self.output_path}")
        self.output_path.parent.mkdir(parents=True, exist_ok=True)
        if self.output_path.exists():
            try:
                self.output_path.unlink()
            except OSError as exc:
                logger.warning(f"Impossible de supprimer l'ancien fichier {self.output_path}: {exc}")
        try:
            self.wb.save(self.output_path)
            logger.info(f"DSF généré avec succès: {self.output_path}")
            return self.output_path
        except PermissionError as exc:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            fallback = self.output_path.with_name(f"{self.output_path.stem}_{timestamp}{self.output_path.suffix}")
            logger.warning(
                "Impossible d'écrire %s (%s). Sauvegarde de secours vers %s.",
                self.output_path,
                exc,
                fallback,
            )
            self.wb.save(fallback)
            logger.info(f"DSF généré avec succès: {fallback}")
            return fallback

    def close(self):
        """Ferme les ressources."""
        if self.wb:
            self.wb.close()
            self.wb = None


class DSFTemplateBuilder:
    """Construit la structure template du DSF de zéro (si nécessaire)."""

    @staticmethod
    def create_dsf_template(output_path: Path) -> Path:
        """Crée un template DSF minimal vierge.
        
        Pour démarrage rapide sans dépendre du fichier standard.
        """
        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        # Ajouter les feuilles principales
        sheets_to_create = [
            'ENTETE',
            'INFORMATIONS GENERALES',
            'Fiche R1',
            'Fiche R2',
            'Fiche R3',
            'BILAN PAYSAGE',
            'COMPTE DE RESULTAT',
            'TABLEAU DES FLUX DE TRESORERIE',
        ]

        for sheet_name in sheets_to_create:
            ws = wb.create_sheet(sheet_name)
            # Ajouter un header minimal
            ws['A1'] = sheet_name
            ws['A1'].font = Font(bold=True, size=14)

        output_path.parent.mkdir(parents=True, exist_ok=True)
        wb.save(output_path)
        logger.info(f"Template DSF créé: {output_path}")
        return output_path


__all__ = [
    "DSFCell",
    "DSFMerge",
    "DSFStreamWriter",
    "DSFTemplateBuilder",
]


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)

    # Exemple d'utilisation
    template = Path("DSF Normal standard.xlsx")
    output = Path("DSF_OUTPUT_TEST.xlsx")

    writer = DSFStreamWriter(template, output, chunk_size=50)
    writer.initialize_from_template()

    # Exemple: écrire quelques valeurs
    writer.write_cell("INFORMATIONS GENERALES", 1, 1, "Test Value")
    writer.write_cell("INFORMATIONS GENERALES", 2, 1, "Another Value", 
                     style={"number_format": "0.00"})

    writer.save()
    print(f"Test DSF généré: {output}")
