"""
MODULE DE GESTION DES CALCULS - DSF MAPPER V7
==============================================

Gère les colonnes nécessitant des calculs:
- SOMMES (TOTAL, SOUS-TOTAL)
- DIFFÉRENCES (Variation = Clôture - Ouverture)
- POURCENTAGES (Taux de variation)
- FORMULES EXCEL (=SOMME(), etc)
"""

import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string

class CalculationManager:
    """Gestionnaire de calculs pour colonnes DSF"""
    
    def __init__(self, ws, column_info):
        """
        Args:
            ws: Worksheet openpyxl
            column_info: Dict {col_letter: col_type} du mapper
        """
        self.ws = ws
        self.column_info = column_info
    
    def is_total_row(self, row_idx):
        """Détecter si une ligne est un TOTAL/SOUS-TOTAL nécessitant une somme
        
        Critères:
        - Libellé contient "TOTAL", "SOUS-TOTAL", "SOMME"
        - Généralement en gras ou style différent
        - NOUVEAU: Vérifier aussi si c'est une ligne vide suivie de texte (séparateur)
        """
        # Vérifier colonnes de labels (A, B, C)
        for col in ['A', 'B', 'C']:
            try:
                cell_value = self.ws[f'{col}{row_idx}'].value
                if cell_value and isinstance(cell_value, str):
                    text_upper = cell_value.upper()
                    # Mots-clés de totalisation
                    if any(keyword in text_upper for keyword in [
                        'TOTAL', 'SOUS-TOTAL', 'SOUS TOTAL',
                        'SOMME', 'CUMUL', 'ENSEMBLE',
                        'GENERAL', 'GLOBAL'  # NOUVEAU
                    ]):
                        return True
                    
                    # NOUVEAU: Détecter les lignes récapitulatives (commence par numéro de section)
                    # Exemple: "3. TOTAL IMMOBILISATIONS", "IV. TOTAL ACTIF"
                    if (text_upper.startswith(('I.', 'II.', 'III.', 'IV.', 'V.', 'VI.')) or
                        text_upper.startswith(tuple(f'{i}.' for i in range(1, 20)))):
                        if 'TOTAL' in text_upper:
                            return True
            except:
                pass
        
        return False
    
    def has_template_formula(self, row_idx, col_letter):
        """Vérifier si la cellule du TEMPLATE a déjà une formule Excel
        
        Returns:
            str: La formule si présente, None sinon
        """
        try:
            cell = self.ws[f'{col_letter}{row_idx}']
            if hasattr(cell, 'value') and isinstance(cell.value, str):
                if cell.value.startswith('='):
                    return cell.value
        except:
            pass
        return None
    
    def parse_complex_formula(self, formula):
        """Parser une formule complexe pour comprendre sa structure
        
        Args:
            formula: Formule Excel (ex: "=D15+E15-F15")
        
        Returns:
            dict: {
                'type': 'complex',
                'references': ['D15', 'E15', 'F15'],
                'operations': ['+', '-'],
                'template': '={ref1}+{ref2}-{ref3}'
            }
        """
        import re
        
        if not formula or not formula.startswith('='):
            return None
        
        # Extraire les références de cellules (ex: D15, AA20)
        cell_pattern = r'([A-Z]{1,3})(\d+)'
        matches = re.findall(cell_pattern, formula)
        
        if not matches:
            return None
        
        references = [f'{col}{row}' for col, row in matches]
        
        # Extraire les opérateurs
        operators = re.findall(r'[+\-*/()]', formula)
        
        return {
            'type': 'complex',
            'references': references,
            'operators': operators,
            'original': formula
        }
    
    def adapt_formula_to_row(self, formula, target_row):
        """Adapter une formule template à une ligne spécifique
        
        Args:
            formula: Formule originale (ex: "=D15+E15-F15")
            target_row: Ligne cible (ex: 20)
        
        Returns:
            str: Formule adaptée (ex: "=D20+E20-F20")
        """
        import re
        
        if not formula or not formula.startswith('='):
            return None
        
        # Remplacer tous les numéros de ligne par target_row
        pattern = r'([A-Z]{1,3})(\d+)'
        
        def replace_row(match):
            col = match.group(1)
            return f'{col}{target_row}'
        
        adapted = re.sub(pattern, replace_row, formula)
        return adapted
    
    def find_sum_range(self, total_row_idx, col_letter):
        """Trouver les lignes à sommer pour calculer un TOTAL
        
        Remonte depuis total_row_idx jusqu'à trouver:
        - Une ligne vide
        - Un autre TOTAL
        - Un header
        
        Returns:
            (start_row, end_row) ou None si pas de range valide
        """
        start_row = None
        
        # Remonter pour trouver le début
        for search_row in range(total_row_idx - 1, max(10, total_row_idx - 30), -1):
            # Vérifier si ligne vide
            cell_val = self.ws[f'{col_letter}{search_row}'].value
            
            # Si vide OU header OU autre total → limite trouvée
            if not cell_val:
                start_row = search_row + 1
                break
            
            if self.is_total_row(search_row):
                start_row = search_row + 1
                break
        
        if not start_row or start_row >= total_row_idx:
            return None
        
        return (start_row, total_row_idx - 1)
    
    def calculate_sum(self, col_letter, start_row, end_row):
        """Calculer la somme d'une plage de cellules
        
        Returns:
            float: Somme calculée
            None: Si pas de valeurs numériques
        """
        total = 0
        has_values = False
        
        for row_idx in range(start_row, end_row + 1):
            cell_val = self.ws[f'{col_letter}{row_idx}'].value
            
            if isinstance(cell_val, (int, float)):
                total += cell_val
                has_values = True
        
        return total if has_values else None
    
    def create_sum_formula(self, col_letter, start_row, end_row):
        """Créer une formule Excel =SOMME(...)
        
        Returns:
            str: Formule Excel (ex: "=SUM(D15:D20)")
        """
        return f"=SUM({col_letter}{start_row}:{col_letter}{end_row})"
    
    def calculate_difference(self, opening_val, closing_val):
        """Calculer une différence (Variation = Closing - Opening)
        
        Returns:
            float: Différence
            None: Si valeurs manquantes
        """
        if opening_val is None or closing_val is None:
            return None
        
        return closing_val - opening_val
    
    def calculate_percentage(self, value, base_value):
        """Calculer un pourcentage (value / base_value * 100)
        
        Returns:
            float: Pourcentage (ex: 15.5 pour 15.5%)
            None: Si division par zéro ou valeurs manquantes
        """
        if base_value is None or value is None or base_value == 0:
            return None
        
        return (value / base_value) * 100
    
    def get_column_for_type(self, col_type):
        """Trouver la lettre de colonne pour un type donné
        
        Args:
            col_type: 'OPENING', 'CLOSING', 'MOVEMENTS', etc
        
        Returns:
            str: Lettre de colonne ou None
        """
        for col_letter, ctype in self.column_info.items():
            if ctype == col_type:
                return col_letter
        return None
    
    def should_calculate_cell(self, row_idx, col_letter, col_type):
        """Déterminer s'il faut calculer cette cellule au lieu de la remplir
        
        Returns:
            tuple: (should_calc, calc_type, extra_info)
            - should_calc: bool - Faut-il calculer?
            - calc_type: str - Type de calcul ('SUM', 'DIFF', 'PCT', 'FORMULA', None)
            - extra_info: dict - Informations supplémentaires (ex: formule template)
        """
        # PRIORITÉ 1: Si cellule template a déjà une FORMULE → la préserver/adapter
        template_formula = self.has_template_formula(row_idx, col_letter)
        if template_formula:
            # Parser la formule pour comprendre sa structure
            formula_info = self.parse_complex_formula(template_formula)
            return (True, 'FORMULA', {'formula': template_formula, 'parsed': formula_info})
        
        # PRIORITÉ 2: Si ligne TOTAL → faire une SOMME
        if self.is_total_row(row_idx):
            if col_type in ['OPENING', 'CLOSING', 'MOVEMENTS', 
                           'CALC_TOTAL_OPENING', 'CALC_TOTAL_MOVEMENTS']:
                return (True, 'SUM', {})
        
        # PRIORITÉ 3: Si colonne RESULTAT/SOLDE → calculer DIFFÉRENCE
        if col_type == 'CALC_RESULT':
            return (True, 'DIFF', {})
        
        # PRIORITÉ 4: Si colonne POURCENTAGE (à détecter dans les headers)
        # TODO: Ajouter détection de colonnes de pourcentage
        
        return (False, None, {})
    
    def apply_calculation(self, row_idx, col_letter, col_type, calc_type, extra_info=None, use_formulas=False):
        """Appliquer le calcul approprié à une cellule
        
        Args:
            row_idx: Numéro de ligne
            col_letter: Lettre de colonne
            col_type: Type de colonne
            calc_type: Type de calcul ('SUM', 'DIFF', 'PCT', 'FORMULA')
            extra_info: Informations supplémentaires (dict)
            use_formulas: Si True, créer formules Excel au lieu de valeurs
        
        Returns:
            value: Valeur calculée ou formule Excel
            None: Si calcul impossible
        """
        extra_info = extra_info or {}
        
        if calc_type == 'FORMULA':
            # Formule complexe du template à adapter
            template_formula = extra_info.get('formula')
            if template_formula:
                # Adapter la formule à la ligne actuelle
                adapted = self.adapt_formula_to_row(template_formula, row_idx)
                return adapted if adapted else template_formula
        
        elif calc_type == 'SUM':
            # Trouver la plage à sommer
            sum_range = self.find_sum_range(row_idx, col_letter)
            
            if not sum_range:
                return None
            
            start_row, end_row = sum_range
            
            if use_formulas:
                # Créer formule Excel
                return self.create_sum_formula(col_letter, start_row, end_row)
            else:
                # Calculer la somme
                return self.calculate_sum(col_letter, start_row, end_row)
        
        elif calc_type == 'DIFF':
            # Calculer différence entre colonnes
            opening_col = self.get_column_for_type('OPENING')
            closing_col = self.get_column_for_type('CLOSING')
            
            if not opening_col or not closing_col:
                return None
            
            opening_val = self.ws[f'{opening_col}{row_idx}'].value
            closing_val = self.ws[f'{closing_col}{row_idx}'].value
            
            if use_formulas:
                return f"={closing_col}{row_idx}-{opening_col}{row_idx}"
            else:
                return self.calculate_difference(opening_val, closing_val)
        
        elif calc_type == 'PCT':
            # Calculer pourcentage
            # TODO: Définir quelle colonne diviser par quelle autre
            pass
        
        return None


def get_calculation_manager(ws, column_info):
    """Factory pour créer un gestionnaire de calculs
    
    Usage:
        calc_mgr = get_calculation_manager(ws, column_info)
        
        # Vérifier si une cellule doit être calculée
        should_calc, calc_type = calc_mgr.should_calculate_cell(row_idx, col_letter, col_type)
        
        if should_calc:
            value = calc_mgr.apply_calculation(row_idx, col_letter, col_type, calc_type)
            if value:
                ws[f'{col_letter}{row_idx}'].value = value
    """
    return CalculationManager(ws, column_info)


# EXEMPLE D'UTILISATION
if __name__ == "__main__":
    # Test avec un fichier DSF
    wb = openpyxl.load_workbook(r"templates\DSF Normal standard.xlsx")
    ws = wb['NOTE 3A']
    
    # Simuler column_info du mapper
    column_info = {
        'D': 'OPENING',
        'E': 'MOVEMENTS',
        'F': 'CLOSING',
    }
    
    calc_mgr = get_calculation_manager(ws, column_info)
    
    # Tester détection de lignes TOTAL
    print("Détection des lignes TOTAL:")
    for row_idx in range(10, 50):
        if calc_mgr.is_total_row(row_idx):
            label = ws[f'A{row_idx}'].value or ws[f'B{row_idx}'].value
            print(f"  Row {row_idx}: {label}")
    
    # Tester calcul de somme
    print("\nTest calcul de somme pour colonne D:")
    for row_idx in range(20, 25):
        should_calc, calc_type = calc_mgr.should_calculate_cell(row_idx, 'D', 'OPENING')
        if should_calc:
            value = calc_mgr.apply_calculation(row_idx, 'D', 'OPENING', calc_type)
            print(f"  Row {row_idx}: Calculé = {value}")
