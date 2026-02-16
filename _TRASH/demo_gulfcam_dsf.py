#!/usr/bin/env python3
"""
Démonstration complète du système DSF avec pré-remplissage GULFCAM SAS
Génère un DSF pré-rempli avec toutes les informations générales
"""

import logging
import sys
from pathlib import Path
from datetime import date

from dsf_general_info import GULFCAM_SAS, format_currency, format_date
from dsf_prefill import DSF_Prefiller, prefill_dsf_with_infos
from dsf_pipeline import DSFPipeline, DSFPipelineConfig

# Configuration logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def print_header(title: str):
    """Affiche un en-tête formaté"""
    print("\n" + "=" * 80)
    print(f"  {title}")
    print("=" * 80 + "\n")

def print_section(title: str):
    """Affiche un titre de section"""
    print(f"\n{'─' * 80}")
    print(f"► {title}")
    print(f"{'─' * 80}\n")

def demo_prefill_entete():
    """Démontre le pré-remplissage de la page ENTETE"""
    print_section("1. PRÉ-REMPLISSAGE PAGE ENTETE")
    
    print(f"Dénomination: {GULFCAM_SAS.denomination_sociale}")
    print(f"Sigle: {GULFCAM_SAS.sigle_usuel}")
    print(f"N° Fiscal: {GULFCAM_SAS.num_identification_fiscale}")
    print(f"Adresse: {GULFCAM_SAS.adresse_complete}")
    print(f"Exercice: {format_date(GULFCAM_SAS.exercice_debut)} au {format_date(GULFCAM_SAS.exercice_fin)}")
    print(f"Système: {GULFCAM_SAS.systeme_normal}")

def demo_r1():
    """Démontre le pré-remplissage de la fiche R1"""
    print_section("2. PRÉ-REMPLISSAGE FICHE R1 (Informations Exercice)")
    
    print(f"Entreprise: {GULFCAM_SAS.denomination_sociale}")
    print(f"Exercice comptable: {format_date(GULFCAM_SAS.exercice_debut)} à {format_date(GULFCAM_SAS.exercice_fin)}")
    print(f"Date d'arrêté: {format_date(GULFCAM_SAS.date_arrete_comptes)}")
    print(f"Durée: {GULFCAM_SAS.duree_mois} mois")

def demo_r2():
    """Démontre le pré-remplissage de la fiche R2"""
    print_section("3. PRÉ-REMPLISSAGE FICHE R2 (Identification & Renseignements)")
    
    print("Informations juridiques:")
    print(f"  Forme juridique: {GULFCAM_SAS.forme_juridique} (SAS)")
    print(f"  Registre fiscal: {GULFCAM_SAS.registre_fiscal} (IUTS)")
    print(f"  Pays siège: {GULFCAM_SAS.pays_siege_social}")
    
    print("\nActivités de l'entreprise:")
    total = 0
    for activite in GULFCAM_SAS.activites:
        print(f"  • {activite.designation}")
        print(f"    Code: {activite.code_nomenclature}")
        print(f"    CA HT: {format_currency(activite.chiffre_affaire_ht)} CFA ({activite.pourcentage_ca}%)")
        total += activite.chiffre_affaire_ht
    
    print(f"\n  TOTAL CA HT: {format_currency(GULFCAM_SAS.chiffre_affaire_total_ht)} CFA")

def demo_r3():
    """Démontre le pré-remplissage de la fiche R3"""
    print_section("4. PRÉ-REMPLISSAGE FICHE R3 (Dirigeants & Conseil)")
    
    print("Dirigeants:")
    for i, dirigeant in enumerate(GULFCAM_SAS.dirigeants, 1):
        print(f"  {i}. {dirigeant.nom} {dirigeant.prenoms}")
        print(f"     Qualité: {dirigeant.qualite}")
        print(f"     Adresse: {dirigeant.adresse}")
    
    print(f"\nConseil d'administration ({len(GULFCAM_SAS.conseil_administration)} membres):")
    for i, membre in enumerate(GULFCAM_SAS.conseil_administration, 1):
        print(f"  {i}. {membre.nom} {membre.prenoms} - {membre.qualite}")

def demo_capital():
    """Démontre le pré-remplissage du capital (Note 13)"""
    print_section("5. PRÉ-REMPLISSAGE NOTE 13 (Structure Capital)")
    
    print("Actionnaires principaux:")
    total_montant = 0
    for actionnaire in GULFCAM_SAS.actionnaires[:5]:
        nom = f"{actionnaire.nom} {actionnaire.prenoms}".strip()
        print(f"  • {nom}")
        print(f"    Actions: {actionnaire.nombre:,} | Montant: {format_currency(actionnaire.montant_total)} CFA")
        total_montant += actionnaire.montant_total
    
    if len(GULFCAM_SAS.actionnaires) > 5:
        print(f"  ... et {len(GULFCAM_SAS.actionnaires) - 5} autres actionnaires")
    
    print(f"\nCapital appelé: {format_currency(GULFCAM_SAS.capital_social_total)} CFA")
    print(f"Capital non appelé: {format_currency(GULFCAM_SAS.capital_non_appele)} CFA")

def generate_prefilled_dsf():
    """Génère le fichier DSF pré-rempli"""
    print_section("6. GÉNÉRATION DU DSF PRÉ-REMPLI")
    
    template_file = Path("DSF Normal standard.xlsx")
    output_file = Path("DSF_GULFCAM_PREFILLED.xlsx")
    
    if not template_file.exists():
        print(f"❌ ERREUR: Template non trouvé: {template_file}")
        print("   Assurez-vous que 'DSF Normal standard.xlsx' est dans le répertoire")
        return False
    
    try:
        print(f"Chargement du template: {template_file}")
        prefiller = DSF_Prefiller(str(template_file))
        prefiller.load_template()
        
        print("Pré-remplissage des pages...")
        prefiller.fill_all(GULFCAM_SAS)
        
        print(f"Sauvegarde du fichier: {output_file}")
        prefiller.save(str(output_file))
        
        print(f"✓ DSF pré-rempli généré avec succès: {output_file}")
        return True
        
    except Exception as e:
        print(f"❌ ERREUR lors de la génération: {e}")
        logger.exception("Erreur de pré-remplissage")
        return False

def generate_full_dsf_with_balance():
    """Génère un DSF complet avec données de balance (simulation)"""
    print_section("7. DSF COMPLET (Données + Balance Complète)")
    
    output_file = Path("DSF_GULFCAM_COMPLET.xlsx")
    
    try:
        config = DSFPipelineConfig(
            template_dsf="DSF Normal standard.xlsx",
            balance_input="DSF_GULFCAM_PREFILLED.xlsx",  # Utiliser le DSF pré-rempli comme balance
            dsf_output=str(output_file),
            exercice_year=2024,
            company_name=GULFCAM_SAS.denomination_sociale,
            company_id=GULFCAM_SAS.num_identification_fiscale
        )
        
        print("Configuration du pipeline DSF:")
        print(f"  Template: {config.template_dsf}")
        print(f"  Sortie: {config.dsf_output}")
        print(f"  Entreprise: {config.company_name}")
        print(f"  N° ID: {config.company_id}")
        
        # Vérification des fichiers
        if not Path(config.template_dsf).exists():
            print(f"⚠ Template manquant: {config.template_dsf}")
            return False
        
        print("✓ Pipeline DSF configuré et prêt")
        return True
        
    except Exception as e:
        print(f"❌ ERREUR: {e}")
        logger.exception("Erreur pipeline")
        return False

def main():
    """Exécute la démonstration complète"""
    print_header("DÉMONSTRATION SYSTÈME DSF - GULFCAM SAS")
    
    # Afficher les infos GULFCAM
    print(f"Entité testée: {GULFCAM_SAS.denomination_sociale}")
    print(f"N° d'identification: {GULFCAM_SAS.num_identification_fiscale}")
    print(f"Exercice: {format_date(GULFCAM_SAS.exercice_debut)} - {format_date(GULFCAM_SAS.exercice_fin)}")
    
    # Exécuter les démos
    demo_prefill_entete()
    demo_r1()
    demo_r2()
    demo_r3()
    demo_capital()
    
    # Générer les fichiers
    success_prefill = generate_prefilled_dsf()
    success_full = generate_full_dsf_with_balance()
    
    # Résumé
    print_header("RÉSUMÉ DE LA DÉMONSTRATION")
    
    print("✓ Données GULFCAM SAS chargées avec succès")
    print(f"  - {len(GULFCAM_SAS.activites)} activités")
    print(f"  - {len(GULFCAM_SAS.dirigeants)} dirigeants")
    print(f"  - {len(GULFCAM_SAS.conseil_administration)} membres du conseil")
    print(f"  - {len(GULFCAM_SAS.actionnaires)} actionnaires")
    
    if success_prefill:
        print("\n✓ DSF pré-rempli généré: DSF_GULFCAM_PREFILLED.xlsx")
    else:
        print("\n✗ Erreur lors de la génération du DSF pré-rempli")
    
    if success_full:
        print("✓ Pipeline DSF configuré et prêt")
    
    print("\n" + "=" * 80)
    print("PROCHAINES ÉTAPES:")
    print("  1. Pré-remplissage des pages générales: DSF_GULFCAM_PREFILLED.xlsx")
    print("  2. Intégration avec la balance complète pour génération DSF complète")
    print("  3. Validation des données et ajustements")
    print("=" * 80 + "\n")

if __name__ == "__main__":
    main()
