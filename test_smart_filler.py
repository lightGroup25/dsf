# -*- coding: utf-8 -*-
"""
Script de test pour le SmartGeneralFiller
Test de reconnaissance automatique des zones de saisie
"""
import json
import logging
from pathlib import Path

# Add src to path
import sys
sys.path.insert(0, str(Path(__file__).parent / "src"))

from smart_general_filler import SmartGeneralFiller
from dsf_general_info import get_gulfcam_config

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s"
)

logger = logging.getLogger(__name__)


def main():
    logger.info("="*70)
    logger.info("TEST SMART GENERAL FILLER - Reconnaissance automatique")
    logger.info("="*70)
    
    # Configuration
    template_path = Path("templates/DSF Normal standard.xlsx")
    output_path = Path("output/test_smart_filled.xlsx")
    report_path = Path("output/test_smart_report.json")
    
    if not template_path.exists():
        logger.error(f"Template not found: {template_path}")
        return
    
    # Obtenir les infos de l'entreprise (GULFCAM par défaut)
    info = get_gulfcam_config()
    
    logger.info(f"\nEntreprise: {info.denomination_sociale}")
    logger.info(f"NIU: {info.num_identification_fiscale}")
    logger.info(f"Exercice: {info.exercice_debut} → {info.exercice_fin}")
    
    # Créer le filler intelligent
    filler = SmartGeneralFiller(template_path)
    
    # Étape 1: Charger et détecter les zones de saisie
    logger.info("\n" + "="*70)
    logger.info("ÉTAPE 1: Détection des zones de saisie")
    logger.info("="*70)
    
    filler.load()
    
    logger.info(f"\n✓ {len(filler.input_zones)} zones de saisie détectées\n")
    
    # Afficher quelques exemples de zones détectées
    logger.info("Exemples de zones détectées:")
    for i, zone in enumerate(filler.input_zones[:20], 1):
        logger.info(f"  {i}. {zone.sheet}!{zone.cell_ref}: '{zone.label}' (type: {zone.zone_type})")
    
    if len(filler.input_zones) > 20:
        logger.info(f"  ... et {len(filler.input_zones) - 20} autres zones")
    
    # Étape 2: Remplir automatiquement
    logger.info("\n" + "="*70)
    logger.info("ÉTAPE 2: Remplissage automatique")
    logger.info("="*70)
    
    filled = filler.fill(info)
    
    logger.info(f"\n✓ {filled} zones remplies avec succès")
    
    # Étape 3: Sauvegarder
    logger.info("\n" + "="*70)
    logger.info("ÉTAPE 3: Sauvegarde")
    logger.info("="*70)
    
    output_path.parent.mkdir(parents=True, exist_ok=True)
    filler.save(output_path)
    logger.info(f"\n✓ Fichier sauvegardé: {output_path}")
    
    # Étape 4: Générer le rapport
    logger.info("\n" + "="*70)
    logger.info("ÉTAPE 4: Rapport de remplissage")
    logger.info("="*70)
    
    report = filler.get_report()
    
    logger.info(f"\nStatistiques globales:")
    logger.info(f"  - Total zones détectées: {report['total_zones']}")
    logger.info(f"  - Total zones remplies: {report['total_filled']}")
    logger.info(f"  - Taux de remplissage: {report['fill_rate']*100:.1f}%")
    
    logger.info(f"\nPar feuille:")
    for sheet, stats in report['by_sheet'].items():
        logger.info(f"  {sheet}:")
        logger.info(f"    - Détectées: {stats['detected']}")
        logger.info(f"    - Remplies: {stats['filled']}")
        logger.info(f"    - Taux: {stats['filled']/stats['detected']*100:.1f}%" if stats['detected'] > 0 else "    - Taux: 0.0%")
    
    # Sauvegarder le rapport JSON
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(json.dumps(report, indent=2, ensure_ascii=False), encoding="utf-8")
    logger.info(f"\n✓ Rapport JSON: {report_path}")
    
    # Afficher quelques exemples de matching
    logger.info("\n" + "="*70)
    logger.info("EXEMPLES DE MATCHING")
    logger.info("="*70)
    
    filled_zones = [z for z in filler.input_zones if z.matched_field][:15]
    for i, zone in enumerate(filled_zones, 1):
        logger.info(f"{i}. {zone.sheet}!{zone.cell_ref}")
        logger.info(f"   Label: '{zone.label}'")
        logger.info(f"   → Champ: {zone.matched_field} (confiance: {zone.confidence:.2f})")
        logger.info("")
    
    # Fermer le workbook
    filler.close()
    
    logger.info("="*70)
    logger.info("✨ TEST TERMINÉ AVEC SUCCÈS")
    logger.info("="*70)
    logger.info(f"\nOuvrez le fichier généré pour vérifier:")
    logger.info(f"  {output_path.absolute()}")


if __name__ == "__main__":
    main()
