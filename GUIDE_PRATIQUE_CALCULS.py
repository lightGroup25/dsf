"""
═══════════════════════════════════════════════════════════════════════════════
GUIDE PRATIQUE - SYSTÈME DE CALCULS POUR FORMULES COMPLEXES
═══════════════════════════════════════════════════════════════════════════════

Ce guide explique comment utiliser le système au quotidien.
"""

GUIDE_COMPLET = """
╔═══════════════════════════════════════════════════════════════════════════════╗
║                     GUIDE PRATIQUE DE GESTION DES CALCULS                     ║
║              Comment gérer les formules complexes dans le DSF                 ║
╚═══════════════════════════════════════════════════════════════════════════════╝

═══════════════════════════════════════════════════════════════════════════════════
1. UTILISATION DE BASE
═══════════════════════════════════════════════════════════════════════════════════

Q1: Je veux que mes TOTALS se calculent automatiquement en Excel?
───────────────────────────────────────────────────────────────────────────────────

✅ C'est DÉJÀ FAIT! Le système génère automatiquement:

    Ligne 16: TOTAL : DOTATIONS
    
    ✅ Colonne I: =SUM(I13:I15)  ← Formule Excel qui se recalcule!

Le jour où quelqu'un modifie I13, I14 ou I15 dans Excel,
la formule I16 se recalculera AUTOMATIQUEMENT.


Q2: Je dois gérer une formule complexe comme =A+B+D-H, comment?
───────────────────────────────────────────────────────────────────────────────────

✅ C'est SUPPORTÉ! Le système:

1. Détecte si le template a une formule
2. Parse la structure: =A15+B15+D15-H15
3. Adapte pour chaque ligne:
   - Ligne 20: =A20+B20+D20-H20
   - Ligne 21: =A21+B21+D21-H21
   - etc.

4. Préserve les références absolues:
   - Template: =A15+B15+$C$10         (C10 est ABSOLUE)
   - Ligne 20: =A20+B20+$C$10         (C10 reste ABSOLUE)


Q3: Comment ça marche pour les POURCENTAGES et VARIATIONS?
───────────────────────────────────────────────────────────────────────────────────

✅ Le système détecte automatiquement!

Si le header contient "VARIATION", "%" ou "TAUX":
   → Crée formule de différence: =Clôture - Ouverture
   
Exemple:
   Header  | OUVERTURE | CLÔTURE | VARIATION
   ────────┼───────────┼─────────┼──────────
   Item 1  | 100       | 120     | =E15-D15
           |           |         | ↓
           |           |         | 20


═══════════════════════════════════════════════════════════════════════════════════
2. ARCHITECTURE INTERNE
═══════════════════════════════════════════════════════════════════════════════════

Le système fonctionne en 4 NIVEAUX DE PRIORITÉ:

┌─────────────────────────────────────────────────────────────────────────┐
│ NIVEAU 1 (PRIORITÉ MAX): Formules Template                             │
├─────────────────────────────────────────────────────────────────────────┤
│ Actions:                                                                │
│   ✓ Détecte si template a formule Excel                                │
│   ✓ Parse la formule (ex: =D15+E15-F15)                                │
│   ✓ Adapte pour chaque ligne                                           │
│   ✓ Préserve références absolues                                       │
│                                                                         │
│ Exemple:                                                                │
│   R15: Client = Ventes + Exports - Remises                             │
│   T15: =D15+E15-F15                                                    │
│   ↓                                                                     │
│   R20: Client = Ventes + Exports - Remises                             │
│   T20: =D20+E20-F20  ✅ (adapté automatiquement)                       │
└─────────────────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────────────────┐
│ NIVEAU 2 (PRIORITÉ HAUTE): Lignes TOTAL                                │
├─────────────────────────────────────────────────────────────────────────┤
│ Actions:                                                                │
│   ✓ Détecte si libellé contient "TOTAL"                                │
│   ✓ Trouve les lignes à sommer (scan haut)                             │
│   ✓ Crée formule =SUM(D10:D14)                                         │
│                                                                         │
│ Exemple:                                                                │
│   R13: Item 1       D13: 100                                            │
│   R14: Item 2       D14: 200                                            │
│   R15: TOTAL        D15: =SUM(D13:D14)  ✅                             │
│                      ↓                                                   │
│                     300 (recalculé automatiquement)                     │
└─────────────────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────────────────┐
│ NIVEAU 3 (PRIORITÉ MOYENNE): Colonnes Calculées                        │
├─────────────────────────────────────────────────────────────────────────┤
│ Détection:                                                              │
│   • Header contient: "VARIATION", "ÉCART", "RÉSULTAT"                  │
│   • Header contient: "TAUX", "%", "POURCENTAGE"                        │
│   • Header contient: "RATIO", "COEFFICIENT"                            │
│                                                                         │
│ Actions:                                                                │
│   ✓ VARIATION: =Colonne_Clôture - Colonne_Ouverture                   │
│   ✓ TAUX: =(Clôture - Ouverture) / Ouverture * 100                     │
│   ✓ RATIO: =Valeur1 / Valeur2                                          │
│                                                                         │
│ Exemple:                                                                │
│   Header | OUVERTUE | CLÔTURE | VARIATION                             │
│   ────────────────────────────────────────                             │
│   Item 1 | 100      | 125     | =D15-C15  ✅ (=25)                    │
└─────────────────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────────────────┐
│ NIVEAU 4 (PRIORITÉ BASSE): Remplissage Défaut                          │
├─────────────────────────────────────────────────────────────────────────┤
│ Si aucun niveau précédent ne s'applique:                               │
│   → Remplir avec matching fuzzy de balance                             │
│   → Sinon laisser vide                                                 │
└─────────────────────────────────────────────────────────────────────────┘


═══════════════════════════════════════════════════════════════════════════════════
3. CAS D'USAGE - TEST INTERACTIF
═══════════════════════════════════════════════════════════════════════════════════

CAS 1: Formule Simple dans Template
──────────────────────────────────────────────────────────────────────────────────

Scénario:
  Template NOTE 3A, cellule D15 contient: =A15+B15

Résultat après mapper:
  D15 (row 1): =A15+B15      ← Inchangée
  D16 (row 2): =A16+B16      ← Adaptée automatiquement
  D17 (row 3): =A17+B17      ← Adaptée automatiquement

Status: ✅ SUCCÈS


CAS 2: Formule Complexe
──────────────────────────────────────────────────────────────────────────────────

Scénario:
  Template NOTE 28, cellule I16 contient: =I13+I14+I15

Résultat:
  I16: =SUM(I13:I15)  ← Converties en SUM automatiquement

Status: ✅ SUCCÈS


CAS 3: Ligne TOTAL sans Données
──────────────────────────────────────────────────────────────────────────────────

Scénario:
  NOTE 3A Row 15: TOTAL
  Mais les lignes au-dessus (13, 14) sont vides

Résultat:
  Pas de formule créée (pas de données à sommer)

Status: ⚠️ ATTENDU (données vides)

Solution:
  Ajouter les données, puis mapper à nouveau


CAS 4: Variation (CLOTURE - OUVERTURE)
──────────────────────────────────────────────────────────────────────────────────

Scénario:
  Header: "OUVERTURE | CLÔTURE | VARIATION"
  Row 15: 100        | 125     | [VIDE]

Résultat:
  Row 15 Variation: =D15-C15  ← Formule créée automatiquement
                   ↓
                   25

Status: ✅ SUCCÈS


═══════════════════════════════════════════════════════════════════════════════════
4. DÉPANNAGE (FAQ)
═══════════════════════════════════════════════════════════════════════════════════

❓ Q: Ma ligne TOTAL n'a pas reçu de formule. Pourquoi?
─────────────────────────────────────────────────────────────────────────────────

Causes possibles:
  1. Les lignes au-dessus ont des données VIDES
     → Mapper ne crée pas de formule si rien à sommer
     → Solution: Ajouter au moins une valeur
  
  2. Le libellé ne contient pas "TOTAL"
     → Mapper ne reconnaît pas la ligne comme TOTAL
     → Solution: Ajouter "TOTAL" au libellé
  
  3. Données avant la ligne TOTAL manquent ou erreurs
     → find_sum_range() ne trouve pas de range valide
     → Solution: Vérifier structure des données

Diagnostic:
  python verify_calculations.py  ← Affiche les TOTAL détectées


❓ Q: Les formules complexes ne s'adaptent pas correctement?
─────────────────────────────────────────────────────────────────────────────────

Causes possibles:
  1. Référence croisée (ex: ='NOTE 3A'!D15)
     → Les références croisées peuvent ne pas s'adapter
     → Solution: Utiliser références locales uniquement
  
  2. Fonctions avancées (ex: INDEX, MATCH)
     → Pas automatiquement adaptées
     → Solution: Mettre à jour manuellement

Diagnostic:
  python demo_formules_complexes.py  ← Affiche types supportés


❓ Q: Comment forcer un calcul spécifique (ex: toujours SOMME)?
─────────────────────────────────────────────────────────────────────────────────

Solution: Modifier le header de colonne
  
  De: "VALEUR SANS %" 
  À:  "TOTAL: VALEUR" ← Force détection comme TOTAL

Ou modifier le code:
  Dans calculation_manager.py, fonction should_calculate_cell()
  Ajouter règles personnalisées pour votre cas


═══════════════════════════════════════════════════════════════════════════════════
5. COMMANDES UTILES
═══════════════════════════════════════════════════════════════════════════════════

# Générer le fichier DSF avec calculs:
python mapper_v7_optimized_formulas.py

# Vérifier les formules générées:
python verify_calculations.py

# Analyser la structure d'une note:
python analyze_note3_issue.py
python GUIDE_CALCULS.py

# Voir les types de formules supportées:
python demo_formules_complexes.py


═══════════════════════════════════════════════════════════════════════════════════
6. EXEMPLE COMPLET - DE TEMPLATE À OUTPUT
═══════════════════════════════════════════════════════════════════════════════════

Template ORIGINAL (NOTE 3A):
──────────────────────────────────────────────────────────────────────────────────
A15: Client Général              D15: 100
A16: Client Spécial              D16: 200
A17: TOTAL : CLIENTS             D17: [FORMULE]

Template a: D15: =... (fonction pour client général)


Balance (données d'import):
──────────────────────────────────────────────────────────────────────────────────
Balance['Client Général'] = 200 (nouvelle valeur)
Balance['Client Spécial'] = 300
...


Fichier OUTPUT (DSF_FINAL_V7_FIXED.xlsx):
──────────────────────────────────────────────────────────────────────────────────
A15: Client Général              D15: 200
     ↑                            ↑
     │ (Libellé du template)      │ (Valeur de balance)
     │                            │
A16: Client Spécial              D16: 300
     ↑                            ↑
     │ (Libellé du template)      │ (Valeur de balance)
     │                            │
A17: TOTAL : CLIENTS             D17: =SUM(D15:D16)
     ↑                            ↑
     │ (Libellé du template)      │ (Formule générée!)
     │                            │
     │                            (500 ← Se recalcule automatiquement)


═══════════════════════════════════════════════════════════════════════════════════
7. DÉCISIONS À PRENDRE
═══════════════════════════════════════════════════════════════════════════════════

Pour NOTE 3A/3B qui n'ont pas de formules actuellement:

OPTION 1: ATTENDRE (RECOMMANDÉ POUR MAINTENANT)
  ✓ Avantage: Pas de changement de code
  ✗ Inconvénient: Les formules n'existent que si données présentes

OPTION 2: CRÉER FORMULES À L'AVANCE
  ✓ Avantage: Prêt pour futures données
  ✓ Avantage: Les formules se calculeront automatiquement
  ✗ Inconvénient: Ajouter logique au mapper
  Implémentation:
    - Générer formules même si données vides
    - Les formules= =SUM(...) se calculeront comme 0 si données vides

OPTION 3: REMPLIR AVEC ZÉROS
  ✓ Avantage: Plus logique pour bilan (0 = pas d'immobilisations)
  ✗ Inconvénient: Peut masquer des données manquantes
  Implémentation:
    - Remplir colonnes data avec 0 par défaut
    - Puis créer formules normalement


═══════════════════════════════════════════════════════════════════════════════════
8. PROCHAINES ÉTAPES
═══════════════════════════════════════════════════════════════════════════════════

✅ COMPLÉTÉ:
  • Système de calculs implémenté
  • Formules complexes supportées
  • Détection améliorée
  • Documentation écrite

⏭️ RECOMMANDÉ:
  1. Tester avec données réelles complètes
  2. Valider formules sur cas métier
  3. Décider pour NOTE 3A/3B (Option 1/2/3)
  4. Documenter règles métier pour calculs spécifiques

╚═══════════════════════════════════════════════════════════════════════════════╝
"""

if __name__ == "__main__":
    print(GUIDE_COMPLET)
