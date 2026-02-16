"""
DÉMONSTRATION - GESTION DES FORMULES COMPLEXES
===============================================

Ce script montre comment le système gère les formules complexes du type A+B+D-H
"""

import openpyxl
from pathlib import Path
from calculation_manager import get_calculation_manager

# ==============================================================================
# 1. PRINCIPE DE GESTION DES FORMULES
# ==============================================================================

PRINCIPE = """
HIÉRARCHIE DE TRAITEMENT (par ordre de priorité):

1️⃣ FORMULES TEMPLATE (PRIORITÉ MAX)
   Si le template a déjà une formule → LA PRÉSERVER et l'ADAPTER
   
   Exemple template:
   - Cellule D15: "=A15+B15-C15"
   
   Action système:
   - Détecte la formule dans le template
   - Parse la structure: references=['A15', 'B15', 'C15'], operators=['+', '-']
   - Adapte pour chaque ligne:
     * Ligne 16: "=A16+B16-C16"
     * Ligne 17: "=A17+B17-C17"
     * etc.

2️⃣ LIGNES TOTAL (PRIORITÉ HAUTE)
   Si libellé contient "TOTAL" → Calculer SOMME
   
   Action système:
   - Trouve le range à sommer (lignes au-dessus)
   - Crée formule: "=SUM(D10:D14)"

3️⃣ COLONNES CALCULÉES (PRIORITÉ MOYENNE)
   Si header contient "VARIATION", "RÉSULTAT" → Calculer DIFFÉRENCE
   
   Action système:
   - Identifie colonnes sources (OUVERTURE, CLÔTURE)
   - Crée formule: "=F15-D15"

4️⃣ REMPLISSAGE BALANCE (PRIORITÉ BASSE)
   Sinon → Remplir avec fuzzy matching depuis balance
"""

# ==============================================================================
# 2. EXEMPLES DE FORMULES SUPPORTÉES
# ==============================================================================

FORMULES_SUPPORTEES = {
    'Simple Addition': {
        'template': '=A15+B15',
        'adaptation_row_20': '=A20+B20',
        'description': 'Addition simple de deux colonnes'
    },
    
    'Soustraction': {
        'template': '=D15-E15',
        'adaptation_row_20': '=D20-E20',
        'description': 'Différence entre colonnes'
    },
    
    'Formule Composée': {
        'template': '=A15+B15+D15-H15',
        'adaptation_row_20': '=A20+B20+D20-H20',
        'description': 'Formule avec plusieurs opérations'
    },
    
    'Avec Multiplication': {
        'template': '=D15*0.5+E15',
        'adaptation_row_20': '=D20*0.5+E20',
        'description': 'Calcul avec coefficient'
    },
    
    'Avec Parenthèses': {
        'template': '=(D15+E15)*0.2',
        'adaptation_row_20': '=(D20+E20)*0.2',
        'description': 'Formule avec priorité d\'opérations'
    },
    
    'Références Absolues': {
        'template': '=A15+$B$10',
        'adaptation_row_20': '=A20+$B$10',
        'description': 'Mélange de références relatives et absolues'
    }
}

# ==============================================================================
# 3. PROCESSUS DE DÉTECTION ET ADAPTATION
# ==============================================================================

PROCESSUS = """
ÉTAPE 1: SCAN DU TEMPLATE
Pour chaque cellule du template DSF:
    - Vérifier si contient formule Excel (commence par '=')
    - Si oui → Extraire et parser la formule

ÉTAPE 2: PARSING DE LA FORMULE
Extraire les composants:
    - Références de cellules: A15, B20, AA100, etc.
    - Opérateurs: +, -, *, /, (, )
    - Constantes: 0.5, 100, etc.
    - Fonctions: SUM(), IF(), etc.

ÉTAPE 3: ADAPTATION PAR LIGNE
Pour chaque ligne de données (15, 16, 17...):
    - Remplacer les numéros de ligne dans les références
    - Préserver les références absolues ($B$10)
    - Garder les constantes inchangées

EXEMPLE COMPLET:
Template ligne 15: =A15+B15*0.5-$C$10
    ↓
Adaptation ligne 20: =A20+B20*0.5-$C$10
Adaptation ligne 25: =A25+B25*0.5-$C$10
"""

# ==============================================================================
# 4. GESTION DES CAS SPÉCIAUX
# ==============================================================================

CAS_SPECIAUX = {
    'Formule avec SUM()': {
        'template': '=SUM(A15:A20)+B15',
        'comportement': 'Préservé tel quel (pas d\'adaptation automatique)',
        'raison': 'Les plages (A15:A20) doivent être gérées manuellement'
    },
    
    'Formule avec IF()': {
        'template': '=IF(A15>0, A15, 0)',
        'comportement': 'Adapté: =IF(A20>0, A20, 0)',
        'raison': 'Les références sont adaptées dans les conditions'
    },
    
    'Références à autres sheets': {
        'template': '=\"NOTE 3A\"!D15+D15',
        'comportement': 'Préservé avec adaptation des refs locales',
        'raison': 'Les noms de sheets sont préservés'
    },
    
    'Cellule fusionnée': {
        'template': 'Cellule fusionnée avec formule',
        'comportement': 'SKIP - Ne pas modifier',
        'raison': 'Protection des zones fusionnées'
    }
}

# ==============================================================================
# 5. DÉMONSTRATION PRATIQUE
# ==============================================================================

def demo_formula_parsing():
    """Démontrer le parsing et l'adaptation de formules"""
    
    print("="*80)
    print("DÉMONSTRATION - PARSING ET ADAPTATION DE FORMULES")
    print("="*80)
    
    # Créer un gestionnaire factice pour la démo
    template_file = Path(r"c:\Users\Emmaneul Ambadiang\Desktop\DSF\templates\DSF Normal standard.xlsx")
    
    if not template_file.exists():
        print(f"Template not found: {template_file}")
        return
    
    wb = openpyxl.load_workbook(template_file)
    
    # Tester sur une sheet avec formules
    test_sheets = ['NOTE 3A', 'NOTE 3B', 'CF2']
    
    for sheet_name in test_sheets:
        if sheet_name not in wb.sheetnames:
            continue
        
        ws = wb[sheet_name]
        
        print(f"\n{'='*80}")
        print(f"SHEET: {sheet_name}")
        print('='*80)
        
        formulas_found = 0
        
        # Scanner pour formules
        for row_idx in range(10, min(50, ws.max_row + 1)):
            for col_idx in range(1, 15):
                col_letter = openpyxl.utils.get_column_letter(col_idx)
                cell = ws[f'{col_letter}{row_idx}']
                
                if hasattr(cell, 'value') and isinstance(cell.value, str):
                    if cell.value.startswith('='):
                        formulas_found += 1
                        
                        if formulas_found <= 3:  # Montrer 3 premiers exemples
                            print(f"\n  📐 FORMULE DÉTECTÉE:")
                            print(f"     Cellule: {col_letter}{row_idx}")
                            print(f"     Formule template: {cell.value}")
                            
                            # Simuler adaptation pour ligne 25
                            import re
                            adapted = re.sub(r'([A-Z]{1,3})(\d+)', 
                                           lambda m: f'{m.group(1)}25', 
                                           cell.value)
                            print(f"     Adaptation row 25: {adapted}")
        
        if formulas_found > 0:
            print(f"\n  ✅ Total formules dans {sheet_name}: {formulas_found}")
        else:
            print(f"\n  ℹ️ Aucune formule Excel détectée dans {sheet_name}")

def demo_formula_types():
    """Montrer les différents types de formules supportées"""
    
    print("\n\n" + "="*80)
    print("TYPES DE FORMULES SUPPORTÉES")
    print("="*80)
    
    for formula_type, info in FORMULES_SUPPORTEES.items():
        print(f"\n📌 {formula_type}")
        print(f"   Template:       {info['template']}")
        print(f"   Adapté row 20:  {info['adaptation_row_20']}")
        print(f"   Description:    {info['description']}")

if __name__ == "__main__":
    print(PRINCIPE)
    print("\n")
    
    demo_formula_types()
    print("\n")
    
    demo_formula_parsing()
    
    print("\n\n" + "="*80)
    print("RÉSUMÉ")
    print("="*80)
    print("""
✅ Le système gère automatiquement:
   1. Formules simples (A+B, D-E)
   2. Formules complexes (A+B+D-H)
   3. Formules avec coefficients (A*0.5+B)
   4. Formules avec parenthèses ((A+B)*C)
   5. Adaptation automatique par ligne
   
🔧 Configuration:
   - use_formulas=True → Crée formules Excel
   - preserve_template_formulas=True → Ne jamais écraser formules existantes
   
📝 Pour forcer un type de calcul:
   - Modifier get_column_type_v7() pour détecter headers spécifiques
   - Ajouter règles dans should_calculate_cell()
    """)
