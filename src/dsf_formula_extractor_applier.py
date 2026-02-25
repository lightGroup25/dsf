# -*- coding: utf-8 -*-
"""
Extracteur / applicateur de formules Excel DSF

Extrait toutes les formules Excel d'un fichier DSF de référence
et les applique au DSF cible (notre remplissage).
"""
from __future__ import annotations

import logging
from pathlib import Path
from typing import Dict, Optional

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

logger = logging.getLogger(__name__)


def extract_formulas_from_dsf(dsf_path: Path) -> Dict[str, Dict[str, str]]:
    """
    Extrait toutes les formules Excel d'un fichier DSF.

    Args:
        dsf_path: Chemin vers le DSF source (template ou DSF rempli)

    Returns:
        {sheet_name: {cell_ref: formula_str}}
    """
    if not dsf_path.exists():
        logger.warning("Fichier DSF référence introuvable: %s", dsf_path)
        return {}

    formulas: Dict[str, Dict[str, str]] = {}
    wb = load_workbook(dsf_path, data_only=False)

    try:
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            sheet_formulas: Dict[str, str] = {}
            for row in ws.iter_rows():
                for cell in row:
                    if cell.value and isinstance(cell.value, str) and cell.value.strip().startswith("="):
                        sheet_formulas[cell.coordinate] = str(cell.value).strip()
            if sheet_formulas:
                formulas[sheet_name] = sheet_formulas
                logger.debug("  %s: %d formules extraites", sheet_name, len(sheet_formulas))
    finally:
        wb.close()

    total = sum(len(f) for f in formulas.values())
    logger.info("Extraction: %d formules depuis %s (%d feuilles)", total, dsf_path.name, len(formulas))
    return formulas


def apply_formulas_to_dsf(
    target_dsf_path: Path,
    formulas: Dict[str, Dict[str, str]],
    *,
    overwrite_existing_formulas: bool = False,
    overwrite_filled_cells: bool = False,
) -> int:
    """
    Applique les formules extraites au DSF cible.

    Args:
        target_dsf_path: Chemin vers le DSF à modifier
        formulas: {sheet_name: {cell_ref: formula_str}}
        overwrite_existing_formulas: Si True, remplace les formules existantes
        overwrite_filled_cells: Si True, écrit aussi dans les cellules avec valeurs numériques

    Returns:
        Nombre de formules appliquées
    """
    if not target_dsf_path.exists():
        logger.warning("DSF cible introuvable: %s", target_dsf_path)
        return 0

    if not formulas:
        logger.warning("Aucune formule à appliquer")
        return 0

    wb = load_workbook(target_dsf_path, data_only=False)
    applied = 0

    try:
        for sheet_name, cell_formulas in formulas.items():
            if sheet_name not in wb.sheetnames:
                logger.debug("Feuille %s absente du DSF cible, ignorée", sheet_name)
                continue

            ws = wb[sheet_name]
            for cell_ref, formula in cell_formulas.items():
                try:
                    cell = ws[cell_ref]
                    existing = cell.value

                    if isinstance(existing, str) and existing.strip().startswith("="):
                        if not overwrite_existing_formulas:
                            continue
                    elif existing is not None and existing != "" and not overwrite_filled_cells:
                        if isinstance(existing, (int, float)) and existing != 0:
                            continue

                    cell.value = formula
                    applied += 1
                except Exception as e:
                    logger.debug("Erreur application %s!%s: %s", sheet_name, cell_ref, e)

        if applied > 0:
            wb.save(target_dsf_path)
            logger.info("✓ %d formules appliquées au DSF: %s", applied, target_dsf_path.name)
    finally:
        wb.close()

    return applied


def extract_and_apply_formulas(
    reference_dsf_path: Path,
    target_dsf_path: Path,
    *,
    overwrite_existing_formulas: bool = False,
    overwrite_filled_cells: bool = False,
) -> int:
    """
    Extrait les formules du DSF référence et les applique au DSF cible.

    Args:
        reference_dsf_path: DSF source (template ou DSF déjà rempli avec formules)
        target_dsf_path: DSF à modifier (notre remplissage)

    Returns:
        Nombre de formules appliquées
    """
    formulas = extract_formulas_from_dsf(reference_dsf_path)
    return apply_formulas_to_dsf(
        target_dsf_path,
        formulas,
        overwrite_existing_formulas=overwrite_existing_formulas,
        overwrite_filled_cells=overwrite_filled_cells,
    )


__all__ = [
    "extract_formulas_from_dsf",
    "apply_formulas_to_dsf",
    "extract_and_apply_formulas",
]


if __name__ == "__main__":
    """
    Usage:
      python dsf_formula_extractor_applier.py [reference_dsf.xlsx] [target_dsf.xlsx]

    Si un seul argument: extrait et applique au même fichier (recopie formules dans cellules vides).
    """
    import sys

    logging.basicConfig(level=logging.INFO, format="%(message)s")

    ref = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("templates/DSF Normal standard.xlsx")
    target = Path(sys.argv[2]) if len(sys.argv) > 2 else Path("output/DSF_OUTPUT_2024.xlsx")

    if not ref.exists():
        print(f"❌ Fichier référence introuvable: {ref}")
        print("   Indiquez un DSF (pas une Balance) comme source de formules.")
        sys.exit(1)
    if not target.exists():
        print(f"❌ DSF cible introuvable: {target}")
        sys.exit(1)

    n = extract_and_apply_formulas(ref, target)
    print(f"\n✓ {n} formules appliquées.")
