# -*- coding: utf-8 -*-
"""
GUIDE PRATIQUE - Utilisation du Système de Calculs
"""

if __name__ == "__main__":
    print("""
╔════════════════════════════════════════════════════════════════════════════╗
║                  GUIDE PRATIQUE - SYSTÈME DE CALCULS                       ║
║                 Gestion des Formules Complexes dans DSF                    ║
╚════════════════════════════════════════════════════════════════════════════╝

1. FONCTIONNEMENT DE BASE:
════════════════════════════════════════════════════════════════════════════

Q: Je veux que mes TOTALS se calculent automatiquement en Excel?
R: C'est déjà fait! Le système génère automatiquement:

   Ligne 16: TOTAL : DOTATIONS
   Colonne I: =SUM(I13:I15)  > Formule Excel qui se recalcule!

Si quelqu'un modifie I13, I14 ou I15 dans Excel,
la formule I16 se recalculera AUTOMATIQUEMENT.


2. FORMULES COMPLEXES:
════════════════════════════════════════════════════════════════════════════

Q: Je dois gérer une formule comme =A+B+D-H, comment?
R: C'est supporté!

   Template ligne 15: =A15+B15+D15-H15
   
   Mapper adapte pour chaque ligne:
   Ligne 16: =A16+B16+D16-H16
   Ligne 17: =A17+B17+D17-H17
   ...

   Les références absolues sont préservées:
   Template: =A15+B15+$C$10
   Ligne 20: =A20+B20+$C$10  (C10 reste absolue)


3. HIÉRARCHIE DE PRIORITÉ:
════════════════════════════════════════════════════════════════════════════

Le système applique les calculs dans cet ordre:

   PRIORITÉ 1: Formules du Template
      - Si le template a une formule Excel
      - La préserver et l'adapter pour chaque ligne
      
   PRIORITÉ 2: Lignes TOTAL
      - Si libellé contient "TOTAL", "SOUS-TOTAL"
      - Créer formule =SUM(...) pour sommer les lignes au-dessus
      
   PRIORITÉ 3: Colonnes Calculées
      - Si header contient "VARIATION", "%", "TAUX"
      - Créer formule de calcul (ex: Clôture - Ouverture)
      
   PRIORITÉ 4: Données de Balance
      - Sinon, remplir avec données du fichier balance


4. TYPES DE CALCULS SUPPORTÉS:
════════════════════════════════════════════════════════════════════════════

Type              Formula                    Détection
────────────────────────────────────────────────────────────────
SUM               =SUM(D15:D20)             "TOTAL", "SOUS-TOTAL"
DIFF              =F15-D15                  "VARIATION", "ÉCART"
PERCENTAGE        =(F15-D15)/D15*100       "%", "TAUX"
RATIO             =D15/D20                  "RATIO"
FORMULA (Complex) =A15+B15*0.5-$C$10       Template a formule


5. EXEMPLE PRATIQUE:
════════════════════════════════════════════════════════════════════════════

Fichier Excel Original (Template):
  Row 15: Item 1                D15: 100
  Row 16: Item 2                D16: 200
  Row 17: TOTAL                 D17: [VIDE]

Après Mapper:
  Row 15: Item 1                D15: 100 (de balance)
  Row 16: Item 2                D16: 200 (de balance)
  Row 17: TOTAL                 D17: =SUM(D15:D16)  <- Formule créée!
                                      = 300 (automatique)

Si balance change:
  + Utilisateur modifie D15 en 500 dans Excel
  + D17 se recalcule automatiquement
  + D17 devient 700


6. CAS D'USAGE:
════════════════════════════════════════════════════════════════════════════

NOTE 28 - DOTATIONS ET PROVISIONS:
  Header: DOTATIONS | 2024 | 2023 | VARIATION
  ───────────────────────────────────────────
  Row 1:  Immeubles | 100  | 80   | =C1-D1 (formule créée!)
  Row 2:  Mobilier  | 50   | 40   | =C2-D2
  Row 15: TOTAL     | 150  | 120  | =SUM(C1:C14) + =D15-E15 (variation!)


BILAN PAYSAGE - ACTIFS:
  Header: LIBELLÉ | VALEUR | COMMENTAIRE
  ───────────────────────────────────────
  Row 1:  Liquidités   | 1000
  ...
  Row 30: TOTAL ACTIF  | =SUM(D15:D29) (formule créée!)


7. DÉPANNAGE:
════════════════════════════════════════════════════════════════════════════

Q: Ma ligne TOTAL n'a pas reçu de formule?
R: Causes possibles:
   1. Pas de données dans les lignes au-dessus (données vides)
   2. Le libellé ne contient pas "TOTAL"
   3. find_sum_range() ne trouve pas de range valide

   Solution: Python verify_calculations.py (affiche les TOTAL détectées)


Q: Les formules complexes ne s'adaptent pas?
R: Causes possibles:
   1. Références croisées (='NOTE 3A'!D15) - limitées
   2. Fonctions avancées (INDEX, MATCH) - pas adaptées

   Solution: Utiliser références locales uniquement


Q: Comment forcer un type de calcul spécifique?
R: Modifier le header de colonne:

   De: "VALEUR"
   À:  "TOTAL: VALEUR"  <- Force détection comme TOTAL

   Ou modifier le code dans calculation_manager.py


8. FICHIERS UTILISATION:
════════════════════════════════════════════════════════════════════════════

# Générer le fichier DSF avec calculs:
  > python mapper_v7_optimized_formulas.py

# Vérifier les formules générées:
  > python verify_calculations.py

# Analyser structure d'une note:
  > python analyze_note3_issue.py

# Voir types de formules supportées:
  > python demo_formules_complexes.py


9. RÉSUMÉ DE L'ARCHITECTURE:
════════════════════════════════════════════════════════════════════════════

calculation_manager.py:
  - CalculationManager: Classe principale
  - is_total_row(): Détecte lignes TOTAL
  - should_calculate_cell(): Détermine si calculer
  - apply_calculation(): Applique le calcul correct
  - parse_complex_formula(): Parse "=A+B-C"
  - adapt_formula_to_row(): Adapte formules

mapper_v7_optimized_formulas.py:
  - Intègre calculation_manager
  - Appelle calc_mgr pour chaque cellule
  - use_formulas=True: Crée formules Excel
  - use_formulas=False: Crée valeurs numériques


10. DÉCISIONS POUR NOTE 3A/3B:
════════════════════════════════════════════════════════════════════════════

Actuellement: NOTE 3A/3B n'ont pas de formules
Raison: Pas de données à sommer

Options:
  1. Attendre les données (option par défaut)
  2. Créer formules à l'avance (=SUM() même si vide)
  3. Remplir avec zéros par défaut

Recommandation: Option 1 pour maintenant
                Passer à Option 2 si données arrivent

────────────────────────────────────────────────────────────────────────────
Résultats actuels:
  NOTE 3A:       0 formules (données vides)
  NOTE 3B:       0 formules (données vides)
  NOTE 28:       2 formules OK
  BILAN PAYSAGE: 1 formule OK
  ────────────
  TOTAL:         3 formules Excel générées

Status: FONCTIONNEL - Prêt pour production avec données complètes
════════════════════════════════════════════════════════════════════════════
    """)
