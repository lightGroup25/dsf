"""
GUIDE COMPLET - GESTION DES COLONNES CALCULÉES
===============================================

Ce document explique comment le système gère les différents types de calculs
dans les colonnes DSF.
"""

# ==============================================================================
# 1. TYPES DE CALCULS SUPPORTÉS
# ==============================================================================

CALCUL_TYPES = {
    'SUM': {
        'description': 'Somme de plusieurs lignes (TOTAL, SOUS-TOTAL)',
        'exemple': 'TOTAL IMMOBILISATIONS = Ligne 15 + Ligne 16 + ... + Ligne 20',
        'formule_excel': '=SUM(D15:D20)',
        'detection': "Libellé contient 'TOTAL', 'SOUS-TOTAL', 'SOMME'",
        'colonnes_concernees': ['OPENING', 'CLOSING', 'MOVEMENTS', 'TOUS_TYPES_NUMERIQUES']
    },
    
    'DIFF': {
        'description': 'Différence entre deux colonnes',
        'exemple': 'VARIATION = CLOTURE - OUVERTURE',
        'formule_excel': '=F15-D15',
        'detection': "Header contient 'VARIATION', 'DIFFERENCE', 'ECART', 'RESULTAT'",
        'colonnes_concernees': ['CALC_RESULT', 'VARIATION']
    },
    
    'PCT': {
        'description': 'Pourcentage / Taux de variation',
        'exemple': 'TAUX VARIATION = (CLOTURE - OUVERTURE) / OUVERTURE * 100',
        'formule_excel': '=(F15-D15)/D15*100',
        'detection': "Header contient '%', 'POURCENTAGE', 'TAUX'",
        'colonnes_concernees': ['CALC_PCT', 'TAUX_VARIATION']
    },
    
    'RATIO': {
        'description': 'Ratio entre deux valeurs',
        'exemple': 'RATIO ENDETTEMENT = DETTES / ACTIF TOTAL',
        'formule_excel': '=D15/D20',
        'detection': "Header contient 'RATIO', 'COEFFICIENT'",
        'colonnes_concernees': ['CALC_RATIO']
    },
    
    'FORMULA': {
        'description': 'Formule Excel complexe détectée dans le template',
        'exemple': '=D15+E15-F15+G15*0.5',
        'formule_excel': 'Préservée du template',
        'detection': 'Cellule contient déjà une formule Excel',
        'colonnes_concernees': ['TOUTES']
    }
}


# ==============================================================================
# 2. EXEMPLES CONCRETS PAR TYPE DE SHEET
# ==============================================================================

EXEMPLES_PAR_SHEET = {
    'NOTE 3A - Immobilisations Incorporelles': {
        'structure': '''
        Ligne 15: Frais de développement      | 1,000,000 | +200,000 | 1,200,000
        Ligne 16: Brevets et licences         | 3,500,000 | +500,000 | 4,000,000
        Ligne 17: Fonds commercial            | 2,000,000 |   +0     | 2,000,000
        Ligne 18: TOTAL IMMOBILISATIONS INCORP | ???      |   ???    |    ???
        ''',
        'calculs_requis': [
            {
                'row': 18,
                'type': 'SUM',
                'col_D': 'OUVERTURE → SUM(D15:D17) = 6,500,000',
                'col_E': 'MOUVEMENTS → SUM(E15:E17) = 700,000',
                'col_F': 'CLOTURE → SUM(F15:F17) = 7,200,000',
            }
        ]
    },
    
    'NOTE 28 - Variation des Provisions': {
        'structure': '''
        Ligne 10: PROVISIONS OUVERTURE        | 73,332,464 | (header)
        Ligne 11: Dotations exercice          |            | +5,000,000
        Ligne 12: Reprises exercice           |            | -2,000,000
        Ligne 13: PROVISIONS CLOTURE          | (à calc)   | (header)
        ''',
        'calculs_requis': [
            {
                'row': 13,
                'type': 'DIFF',
                'description': 'CLOTURE = OUVERTURE + Dotations - Reprises',
                'formule': '=B10+B11-B12',
            }
        ]
    },
    
    'BILAN - Taux de variation': {
        'structure': '''
        Ligne 20: Actif Total       | 100,000  | 120,000  | +20%
        Ligne 21: Passif Total      |  80,000  | 100,000  | +25%
        ''',
        'calculs_requis': [
            {
                'row': 20,
                'col_G': 'TAUX_VARIATION',
                'type': 'PCT',
                'formule': '=(E20-D20)/D20*100',
                'resultat': '20%',
            }
        ]
    }
}


# ==============================================================================
# 3. ALGORITHME DE DÉTECTION ET CALCUL
# ==============================================================================

ALGORITHME_DETECTION = """
Pour chaque cellule (row, col):

ÉTAPE 1: Déterminer le type de colonne
    - Analyser headers (lignes 1-20)
    - Identifier: OPENING, CLOSING, MOVEMENTS, VARIATION, TAUX, etc.

ÉTAPE 2: Vérifier si la ligne nécessite un calcul
    - Est-ce une ligne TOTAL/SOUS-TOTAL ?
    - Le libellé contient "TOTAL", "SOMME", "CUMUL" ?
    
ÉTAPE 3: Identifier le type de calcul requis
    
    SI ligne_est_total ET colonne_est_numerique:
        → Calcul = SUM
        → Trouver range (lignes au-dessus jusqu'au dernier total/vide)
        → Appliquer: SUM(col+start_row:col+end_row)
    
    SINON SI colonne_type == 'VARIATION' OU 'RESULTAT':
        → Calcul = DIFF
        → Trouver colonne_ouverture et colonne_cloture
        → Appliquer: cloture_value - ouverture_value
    
    SINON SI colonne_type == 'TAUX' OU contains '%':
        → Calcul = PCT
        → Identifier base (généralement OUVERTURE)
        → Appliquer: (nouvelle - ancienne) / ancienne * 100
    
    SINON SI cellule_template_a_formule:
        → Calcul = FORMULA
        → Préserver ou adapter la formule existante
    
    SINON:
        → Pas de calcul
        → Remplir avec valeur directe de la balance

ÉTAPE 4: Appliquer le calcul ou remplir
    
    SI calcul_requis:
        - Exécuter le calcul
        - Écrire la valeur OU la formule Excel
    SINON:
        - Chercher valeur dans balance via fuzzy matching
        - Écrire la valeur directe
"""


# ==============================================================================
# 4. CONFIGURATION DES PRIORITÉS
# ==============================================================================

PRIORITES_CALCUL = """
Ordre de priorité pour résoudre les conflits:

1. FORMULES EXCEL EXISTANTES (priorité MAX)
   → Si template contient déjà une formule, LA PRÉSERVER
   
2. LIGNES TOTAL (priorité HAUTE)
   → Toujours calculer les totaux, jamais remplir avec balance
   
3. COLONNES CALCULÉES (priorité MOYENNE)
   → VARIATION, TAUX, RATIO → toujours calculer
   
4. VALEURS DIRECTES (priorité BASSE)
   → Lignes de détail → remplir depuis balance

RÈGLES SPÉCIALES:
- Si cellule non vide dans template → NE PAS MODIFIER
- Si calcul impossible (données manquantes) → LAISSER VIDE
- Si ambiguïté → Privilégier FORMULE EXCEL sur valeur calculée
"""


# ==============================================================================
# 5. INTÉGRATION DANS LE MAPPER V7
# ==============================================================================

CODE_INTEGRATION = """
# Dans process_sheet_v7(), AVANT la boucle de remplissage:

from calculation_manager import get_calculation_manager

# Créer gestionnaire de calculs
calc_mgr = get_calculation_manager(ws, column_info)

# Dans la boucle de remplissage:
for row_idx in range(10, min(300, ws.max_row + 1)):
    for col_letter, col_type in column_info.items():
        cell = ws[f'{col_letter}{row_idx}']
        
        # NOUVEAU: Vérifier si calcul requis
        should_calc, calc_type = calc_mgr.should_calculate_cell(
            row_idx, col_letter, col_type
        )
        
        if should_calc:
            # CALCULER au lieu de remplir depuis balance
            value = calc_mgr.apply_calculation(
                row_idx, col_letter, col_type, calc_type,
                use_formulas=True  # True = formules Excel, False = valeurs
            )
            
            if value:
                cell.value = value
                cells_filled += 1
        else:
            # REMPLIR depuis balance (code existant)
            best_account, score = find_best_account(line_label, col_type, accounts)
            # ... reste du code existant
"""


# ==============================================================================
# 6. OPTIONS DE CONFIGURATION
# ==============================================================================

CONFIG_OPTIONS = {
    'use_excel_formulas': {
        'default': True,
        'description': 'Créer formules Excel au lieu de valeurs calculées',
        'avantage': 'Les formules se recalculent automatiquement si données changent',
        'inconvenient': 'Plus complexe à debugger',
    },
    
    'calculate_totals_automatically': {
        'default': True,
        'description': 'Détecter et calculer automatiquement les lignes TOTAL',
        'avantage': 'Remplis les totaux même si pas dans balance',
        'inconvenient': 'Peut créer des totaux incorrects si détection erronée',
    },
    
    'preserve_template_formulas': {
        'default': True,
        'description': 'Ne jamais écraser les formules existantes du template',
        'avantage': 'Garantit que structure validée est préservée',
        'inconvenient': 'Ne peut pas corriger formules erronées du template',
    },
    
    'min_rows_for_sum': {
        'default': 2,
        'description': 'Nombre minimum de lignes pour créer une somme',
        'exemple': 'Si seulement 1 ligne avant TOTAL, copier la valeur au lieu de SUM()',
    }
}


# ==============================================================================
# 7. TESTS ET VALIDATION
# ==============================================================================

if __name__ == "__main__":
    print("="*80)
    print("GUIDE GESTION DES COLONNES CALCULÉES")
    print("="*80)
    
    print("\n1. TYPES DE CALCULS SUPPORTÉS:")
    print("-" * 80)
    for calc_type, info in CALCUL_TYPES.items():
        print(f"\n  [{calc_type}] {info['description']}")
        print(f"    Exemple: {info['exemple']}")
        print(f"    Formule: {info['formule_excel']}")
        print(f"    Détection: {info['detection']}")
    
    print("\n\n2. EXEMPLES PAR TYPE DE SHEET:")
    print("-" * 80)
    for sheet_name, info in EXEMPLES_PAR_SHEET.items():
        print(f"\n  {sheet_name}:")
        print(info['structure'])
        print(f"    Calculs requis: {len(info['calculs_requis'])}")
    
    print("\n\n3. CONFIGURATION:")
    print("-" * 80)
    for opt_name, opt_info in CONFIG_OPTIONS.items():
        print(f"\n  {opt_name}:")
        print(f"    Default: {opt_info['default']}")
        print(f"    {opt_info['description']}")
    
    print("\n\n" + "="*80)
    print("Pour intégrer: Voir CODE_INTEGRATION dans ce fichier")
    print("="*80)
