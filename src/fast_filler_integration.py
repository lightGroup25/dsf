#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Pipeline Integration - Fast NOTE 20 & 34 Fillers
Branche les remplisseurs structuré avant le fuzzy matching
"""

import logging
import time
from pathlib import Path
from typing import Dict, Optional

from openpyxl import load_workbook, Workbook

logger = logging.getLogger(__name__)


def apply_fast_fillers_pre_semantic(
    template_path: Path,
    accounts: Dict[str, any],  # balance_accounts from normalizer (N)
    accounts_n1: Optional[Dict[str, any]] = None,
    apply_note20: bool = True,
    apply_note34: bool = True,
    apply_extra_notes: bool = True,
) -> Optional[Workbook]:
    """
    Applique les remplisseurs structurés AVANT le fuzzy matching sémantique
    
    Cela permet d'éviter le coût du fuzzy matching pour NOTE 20 et 34.
    
    Args:
        template_path: Path du template DSF
        accounts: Dict[compte] = NormalizedBalanceRow (chargé depuis balance)
        apply_note20: Utiliser fast filler pour NOTE 20
        apply_note34: Utiliser fast filler pour NOTE 34
    
    Returns:
        Workbook avec NOTE 20 et 34 pré-remplies, ou None si pas appliquées
    """
    
    if not apply_note20 and not apply_note34:
        logger.debug("Fast fillers disabled, skipping")
        return None
    
    # Import des fast fillers
    try:
        from note20_34_fast_filler import Note20FastFiller, Note34FastFiller
        from note_rules_fast_filler import NoteRulesFastFiller
    except ImportError:
        logger.error("note20_34_fast_filler module not found, fast fillers disabled")
        return None
    
    start_time = time.time()
    
    # Charge le workbook
    logger.info("Loading template for fast filler pre-processing...")
    wb = load_workbook(template_path, data_only=False)
    
    results = {}
    
    # NOTE 20
    if apply_note20:
        if "NOTE 20" in wb.sheetnames:
            try:
                note20_start = time.time()
                filler20 = Note20FastFiller(wb)
                count = filler20.fill_note20(accounts)
                elapsed = time.time() - note20_start
                results['note20'] = {'count': count, 'time': elapsed}
                logger.info(f"[FAST FILLER] NOTE 20: {count} comptes assignés en {elapsed:.2f}s")
            except Exception as e:
                logger.warning(f"[FAST FILLER] NOTE 20 failed: {e}, falling back to semantic")
                results['note20'] = {'count': 0, 'time': 0, 'error': str(e)}
        else:
            logger.debug("NOTE 20 not found in template")
    
    # NOTE 34
    if apply_note34:
        if "NOTE 34" in wb.sheetnames:
            try:
                note34_start = time.time()
                filler34 = Note34FastFiller(wb)
                count = filler34.fill_note34(accounts)
                elapsed = time.time() - note34_start
                results['note34'] = {'count': count, 'time': elapsed}
                logger.info(f"[FAST FILLER] NOTE 34: {count} assignés en {elapsed:.2f}s")
            except Exception as e:
                logger.warning(f"[FAST FILLER] NOTE 34 failed: {e}, falling back to semantic")
                results['note34'] = {'count': 0, 'time': 0, 'error': str(e)}
        else:
            logger.debug("NOTE 34 not found in template")

    # Extra notes (NOTE 19, 15A, 15B, 16A, C1-NOTE 17, C1-NOTE 25, C2-NOTE 25)
    if apply_extra_notes:
        extra_sheets = [
            "NOTE 18",
            "NOTE 19",
            "NOTE 27A",
            "NOTE 28",
            "C1-NOTE 28",
            "C2-NOTE 28",
            "NOTE 30",
            "NOTE 32",
            "NOTE 15A",
            "NOTE 15B",
            "NOTE 16A",
            "C1-NOTE 17",
            "C1-NOTE 25",
            "C2-NOTE 25",
        ]
        note_filler = NoteRulesFastFiller(wb, accounts, accounts_n1)
        for sheet in extra_sheets:
            if sheet in wb.sheetnames:
                try:
                    count = note_filler.fill_note(sheet)
                    results[sheet] = {"count": count}
                except Exception as e:
                    logger.warning(f"[FAST FILLER] {sheet} failed: {e}, falling back to semantic")
            else:
                logger.debug("%s not found in template", sheet)
    
    total_time = time.time() - start_time
    total_assigned = sum(r.get('count', 0) for r in results.values())
    
    logger.info(f"[FAST FILLER] Total: {total_assigned} assignations en {total_time:.2f}s")
    
    return wb


def should_skip_sheet_in_semantic_filler(sheet_name: str, fast_filler_applied: bool) -> bool:
    """
    Détermine si une feuille doit être zappée dans le fuzzy matching sémantique
    car elle a déjà été remplie par les fast fillers
    
    Args:
        sheet_name: Nom de la feuille
        fast_filler_applied: Si des fast fillers ont été appliqués
    
    Returns:
        True si la feuille peut être zappée
    """
    
    if not fast_filler_applied:
        return False
    
    # Si NOTE 20 ou 34 ont été remplies structurée, on peut largement les zapper
    # du fuzzy matching (gain de temps énorme)
    SKIP_IN_SEMANTIC = {
        "NOTE 18",
        "NOTE 19",
        "NOTE 20",
        "NOTE 27A",
        "NOTE 28",
        "C1-NOTE 28",
        "C2-NOTE 28",
        "NOTE 30",
        "NOTE 32",
        "NOTE 15A",
        "NOTE 15B",
        "NOTE 16A",
        "C1-NOTE 17",
        "C1-NOTE 25",
        "C2-NOTE 25",
        "NOTE 34",
    }
    
    return sheet_name in SKIP_IN_SEMANTIC


# ============================================================================
# CONFIGURATION RECOMMANDÉE PIPELINE
# ============================================================================

FAST_FILLER_CONFIG = {
    "enabled": True,
    "note20": True,  # Remplir NOTE 20 structurée
    "note34": True,  # NOTE 34 aussi si possible
    "skip_in_semantic": True,  # Zapper du fuzzy matching après remplissage
    "fallback_on_error": True,  # Si erreur, utiliser fuzzy normal
}


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format='%(message)s')
    print("[OK] Fast filler integration module loaded")
    print(f"    Default config: {FAST_FILLER_CONFIG}")
