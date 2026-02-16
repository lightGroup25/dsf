# -*- coding: utf-8 -*-
"""
Test complet du système de détection et application de formules de colonnes
Intégré au remplissage sémantique
"""
import json
import logging
from pathlib import Path

import sys
sys.path.insert(0, str(Path(__file__).parent / "src"))

from semantic_balance_filler import SemanticBalanceFiller
from dsf_inventory import DSFInventory
from dsf_general_info import get_gulfcam_config
from smart_general_filler import SmartGeneralFiller

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)

logger = logging.getLogger(__name__)


def main():
    logger.info("="*80)
    logger.info("TEST COMPLET - Formules de colonnes + Remplissage sémantique")
    logger.info("="*80)
    
    # Chemins
    template_path = Path("templates/DSF Normal standard.xlsx")
    balance_path = Path("input/BALANCE AU 31 12 2024 VERSION DU 26 05 25.xlsx")
    inventory_path = Path("data/dsf_inventory.json")
    output_prefilled = Path("output/test_formulas_prefilled.xlsx")
    output_final = Path("output/test_formulas_filled.xlsx")
    
    # Vérifications
    if not template_path.exists():
        logger.error(f"Template not found: {template_path}")
        return
    
    if not balance_path.exists():
        logger.error(f"Balance not found: {balance_path}")
        return
    
    if not inventory_path.exists():
        logger.error(f"Inventory not found: {inventory_path}")
        return
    
    # Charger l'inventaire
    inventory = DSFInventory.from_json(inventory_path)
    
    # Infos entreprise
    info = get_gulfcam_config()
    
    logger.info("\n" + "="*80)
    logger.info("ÉTAPE 1: Pré-remplissage intelligent (ENTÊTE/R1/R2/R3/NOTE13)")
    logger.info("="*80)
    
    smart_filler = SmartGeneralFiller(template_path)
    smart_filler.load()
    filled_general = smart_filler.fill(info)
    smart_filler.save(output_prefilled)
    smart_filler.close()
    
    logger.info(f"✓ {filled_general} zones générales remplies")
    logger.info(f"✓ Template pré-rempli: {output_prefilled}")
    
    logger.info("\n" + "="*80)
    logger.info("ÉTAPE 2: Remplissage sémantique avec détection formules")
    logger.info("="*80)
    
    # Créer le filler sémantique
    filler = SemanticBalanceFiller(
        output_prefilled,
        balance_path,
        inventory,
        fuzzy_threshold=0.6
    )
    
    # Charger (détecte les formules dans les en-têtes)
    logger.info("\nChargement et détection des formules...")
    filler.load()
    
    # Afficher les formules détectées
    if filler.formula_detectors:
        logger.info(f"\n✓ Formules détectées dans {len(filler.formula_detectors)} feuille(s):")
        for sheet_name, detector in filler.formula_detectors.items():
            formulas = detector.get_formula_columns()
            logger.info(f"  {sheet_name}: {len(formulas)} colonne(s) avec formules")
            for col_letter in formulas:
                formula_obj = detector.column_formulas[col_letter]
                logger.info(f"    - Colonne {col_letter}: {formula_obj.header_text}")
                logger.info(f"      Pattern: {formula_obj.formula_pattern}")
    else:
        logger.info("ℹ  Aucune formule détectée dans les en-têtes")
    
    # Remplir (skip les colonnes avec formules)
    logger.info("\nRemplissage sémantique...")
    filled_cells = filler.fill()
    
    logger.info(f"\n✓ {filled_cells} cellules remplies")
    logger.info(f"✓ {filler.formulas_applied} formules de colonnes appliquées")
    
    # Sauvegarder
    logger.info("\nSauvegarde...")
    filler.save(output_final)
    
    logger.info(f"\n✓ Fichier final: {output_final}")
    
    # Statistiques
    logger.info("\n" + "="*80)
    logger.info("STATISTIQUES")
    logger.info("="*80)
    
    stats = filler.get_stats()
    logger.info(f"Cellules assignées: {stats['assignments']}")
    logger.info(f"Cellules non matchées: {stats['unmatched']}")
    logger.info(f"Comptes utilisés: {stats['accounts_used']}")
    logger.info(f"Montant total: {stats['total_amount']:,.2f} FCFA")
    logger.info(f"Formules de colonnes: {filler.formulas_applied}")
    
    logger.info("\n" + "="*80)
    logger.info("✨ TEST TERMINÉ AVEC SUCCÈS")
    logger.info("="*80)
    logger.info(f"\nOuvrez le fichier pour vérifier les formules:")
    logger.info(f"  {output_final.absolute()}")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        logger.error(f"Error: {e}", exc_info=True)
        raise
