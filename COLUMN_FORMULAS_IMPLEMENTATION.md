# SYSTÈME DE DÉTECTION ET APPLICATION DES FORMULES DE COLONNES
## Implémentation complète et intégrée au pipeline DSF

---

## 🎯 Fonctionnalité Implémentée

Le système détecte et implémente automatiquement les **formules définies dans les en-têtes de colonnes** du template DSF.

### Exemple de Détection

**En-tête de colonne :**
```
VARIATION = Clôture - Ouverture
```

**Résultat :** Le système génère automatiquement pour chaque ligne de données :
```excel
Ligne 15: =E15-D15
Ligne 16: =E16-D16
Ligne 17: =E17-D17
...
```

---

## 📦 Fichiers Créés

### 1. **src/column_formula_detector.py** (400+ lignes)

Module de détection de formules dans les en-têtes.

**Classes principales :**

```python
class ColumnFormula:
    """Représente une formule détectée"""
    sheet: str
    column_letter: str
    header_text: str
    formula_pattern: str  # Ex: "{D}-{E}"
    referenced_columns: List[str]  # Ex: ["D", "E"]
    operation_type: str  # "subtraction", "addition", etc.
    
    def generate_formula(self, row_number: int) -> str:
        """Génère =D15-E15 pour la ligne 15"""

class ColumnFormulaDetector:
    """Détecte les formules dans les en-têtes"""
    def detect() -> Dict[str, ColumnFormula]
    def has_formula(col_letter: str) -> bool
    def apply_formula_to_cell(col_letter, row) -> str
```

**Patterns reconnus automatiquement :**

| Pattern dans en-tête | Type | Formule générée |
|---------------------|------|-----------------|
| `= A - B` | subtraction | `={A}-{B}` |
| `= A + B` | addition | `={A}+{B}` |
| `= A * B` | multiplication | `={A}*{B}` |
| `= A / B` | division | `={A}/{B}` |
| `= (A/B)*100` | percentage | `=({A}/{B})*100` |

**Résolution intelligente des références :**

1. **Lettre de colonne directe** : "A", "B", "AA" → utilisé tel quel
2. **Nom complet** : "Clôture", "Ouverture" → recherche dans labels
3. **Fuzzy matching** : "Solde final" → match avec "Clôture" (60% similarité)

---

### 2. **Modifications : src/semantic_balance_filler.py**

**Ajouts :**

```python
from column_formula_detector import ColumnFormulaDetector

class SemanticBalanceFiller:
    def __init__(...):
        # ...
        self.formula_detectors: Dict[str, ColumnFormulaDetector] = {}
        self.formulas_applied: int = 0
```

**Méthode load() - Détection des formules :**

```python
def load(self):
    self.wb = load_workbook(...)
    
    # NOUVEAU : Détecte les formules dans tous les en-têtes
    logger.info("Detecting column formulas in template...")
    for sheet_name in self.wb.sheetnames:
        detector = ColumnFormulaDetector(self.wb[sheet_name])
        formulas = detector.detect()
        if formulas:
            self.formula_detectors[sheet_name] = detector
            logger.info(f"  {sheet_name}: {len(formulas)} formula columns")
    
    # ... reste du chargement balance ...
```

**Méthode fill() - Skip des colonnes avec formules :**

```python
def fill(self):
    for sheet_name in self.wb.sheetnames:
        # Obtenir le détecteur de formules pour cette feuille
        formula_detector = self.formula_detectors.get(sheet_name)
        
        for row in data_rows:
            for col in columns:
                col_letter = get_column_letter(col)
                
                # NOUVEAU : Skip si colonne a formule
                if formula_detector and formula_detector.has_formula(col_letter):
                    logger.debug(f"Skipping {cell} - formula column")
                    continue
                
                # Remplissage normal pour les autres...
    
    # NOUVEAU : Application des formules après remplissage
    logger.info("Applying column formulas...")
    self._apply_column_formulas()
    logger.info(f"Applied {self.formulas_applied} column formulas")
```

**Nouvelle méthode _apply_column_formulas() :**

```python
def _apply_column_formulas(self):
    """Applique les formules Excel dans les colonnes détectées"""
    for sheet_name, detector in self.formula_detectors.items():
        ws = self.wb[sheet_name]
        formula_columns = detector.get_formula_columns()
        
        # Détermine zone de données
        data_start = detector.header_row_detected + 1
        
        for row_idx in range(data_start, ws.max_row + 1):
            # Skip headers/totals
            first_cell = ws.cell(row=row_idx, column=1)
            if 'total' in str(first_cell.value).lower():
                continue
            
            # Applique formule à chaque colonne
            for col_letter in formula_columns:
                cell = ws[f"{col_letter}{row_idx}"]
                
                # Seulement si vide ou numérique
                if cell.value is None or isinstance(cell.value, (int, float)):
                    formula = detector.apply_formula_to_cell(col_letter, row_idx)
                    if formula:
                        cell.value = formula  # Ex: "=E15-D15"
                        self.formulas_applied += 1
```

---

## 🔄 Flux d'Exécution Complet

```
1. Template DSF chargé
   ↓
2. ✓ SmartGeneralFiller remplit ENTÊTE/R1/R2/R3/NOTE13
   ↓
3. ✓ ColumnFormulaDetector scanne tous les en-têtes
   ↓ 
4. ✓ Détecte formules (ex: "VARIATION = Clôture - Ouverture")
   ↓
5. ✓ SemanticBalanceFiller charge balance
   ↓
6. ✓ Remplissage cellule par cellule (SKIP colonnes avec formules)
   ↓
7. ✓ Application formules colonnes (génère =E15-D15, etc.)
   ↓
8. ✓ DSFCalculationApplier ajoute TOTAL/VARIATION/RATIOS
   ↓
9. ✓ Sauvegarde DSF final
   ↓
10. DSF complet avec formules dynamiques
```

---

## 🧪 Tests Créés

### 1. **test_column_formulas.py**
Test standalone du détecteur de formules

```bash
python test_column_formulas.py
```

**Sortie attendue :**
```
Sheet: ACTIF
  ✓ 2 formula column(s) detected:
  Column F:
    Header: 'VARIATION = E - D'
    Pattern: {E}-{D}
    Examples:
      Row 15: =E15-D15
      Row 20: =E20-D20
```

### 2. **test_complete_with_formulas.py**
Test intégration complète

```bash
python test_complete_with_formulas.py
```

**Processus :**
1. Pré-remplissage intelligent (ENTÊTE/R1/R2/R3)
2. Détection formules colonnes
3. Remplissage sémantique balance
4. Application formules détectées
5. Génération DSF final

**Sortie attendue :**
```
✓ 38 zones générales remplies
✓ Formules détectées dans 5 feuille(s)
  ACTIF: 2 colonne(s) avec formules
  PASSIF: 2 colonne(s) avec formules
✓ 1234 cellules remplies
✓ 456 formules de colonnes appliquées
```

---

## 📊 Exemples de Formules Détectées

### Cas 1 : Variation simple
```
En-tête : "VARIATION = Clôture - Ouverture"
Détection : subtraction, colonnes [E, D]
Formule : =E{row}-D{row}
```

### Cas 2 : Total
```
En-tête : "TOTAL = Débit + Crédit"
Détection : addition, colonnes [B, C]
Formule : =B{row}+C{row}
```

### Cas 3 : Pourcentage
```
En-tête : "% = (Valeur / Base) * 100"
Détection : percentage, colonnes [D, E]
Formule : =(D{row}/E{row})*100
```

### Cas 4 : Ratio
```
En-tête : "RATIO = A / B"
Détection : division, colonnes [F, G]
Formule : =F{row}/G{row}
```

---

## ⚙️ Configuration

### Activer (par défaut)
Le système est automatiquement activé lors du remplissage sémantique.

```python
from semantic_balance_filler import SemanticBalanceFiller

filler = SemanticBalanceFiller(template, balance, inventory)
filler.load()  # Détecte automatiquement les formules
filler.fill()  # Skip colonnes formules + applique formules
```

### Vérifier les formules détectées

```python
# Après load()
if filler.formula_detectors:
    for sheet, detector in filler.formula_detectors.items():
        print(f"{sheet}: {len(detector.get_formula_columns())} formulas")
        for col in detector.get_formula_columns():
            formula = detector.column_formulas[col]
            print(f"  {col}: {formula.header_text}")
```

### Statistiques après remplissage

```python
# Après fill()
print(f"Formulas applied: {filler.formulas_applied}")
stats = filler.get_stats()
print(f"Total cells filled: {stats['assignments']}")
```

---

## 🎯 Avantages

| Avant | Après |
|-------|-------|
| Colonnes calculées remplies avec valeurs fixes | Formules Excel dynamiques |
| Pas de recalcul automatique | Recalcul auto si données changent |
| Maintenance manuelle | Automatique |
| Template rigide | Adaptatif |
| Formules perdues au remplissage | Formules préservées et générées |

---

## 🔍 Détection Intelligente

### Stratégie de résolution de références

```python
# Référence "Clôture" dans formule
1. Cherche colonne nommée exactement "Clôture" → Trouvé colonne E
2. Si non trouvé, cherche "cloture" (sans accent) → Trouvé
3. Si non trouvé, fuzzy match "Clôture finale" (60%+) → Trouvé
4. Si non trouvé → Skip formule

# Référence "E" dans formule
1. C'est déjà une lettre → Colonne E directe
```

### Détection ligne d'en-tête

Le système cherche automatiquement la ligne contenant les mots-clés :
- ouverture, opening, clôture, closing
- débit, debit, crédit, credit
- variation, écart, mouvement
- libellé, label, compte, account
- total, montant, valeur

**Seuil :** Au moins 2 mots-clés dans la ligne → C'est l'en-tête

---

## 📝 Logs Générés

```
INFO - Detecting column formulas in template...
INFO -   ACTIF: 2 formula columns detected
INFO -   PASSIF: 2 formula columns detected
DEBUG - Skipping D15 - formula column
DEBUG - Skipping E15 - formula column
INFO - Applying column formulas...
INFO -   Sheet ACTIF: applying formulas to columns D, E
DEBUG -     D15: =E15-C15
DEBUG -     D16: =E16-C16
INFO - Applied 234 column formulas
```

---

## 🚀 Utilisation dans le Pipeline

**Le système est déjà intégré** dans `dsf_pipeline.py` :

```python
# Le pipeline appelle automatiquement
prefilled = self._prefill_general_sections()  # SmartGeneralFiller
dsf_output = self._fill_semantic_balance(prefilled)  # Avec formules colonnes
self._apply_calculations(dsf_output)  # Calculs TOTAL/VARIATION supplémentaires
```

**Rien à configurer !** Tout fonctionne automatiquement.

---

## ✅ Statut

| Composant | Status |
|-----------|--------|
| ColumnFormulaDetector | ✅ Créé |
| Integration SemanticBalanceFiller | ✅ Complète |
| Détection patterns | ✅ 5 types |
| Résolution références | ✅ 3 stratégies |
| Application formules | ✅ Automatique |
| Tests unitaires | ✅ Créés |
| Tests intégration | ✅ Créés |
| Documentation | ✅ Complète |

---

## 🔧 Maintenance

### Ajouter un nouveau pattern de formule

Dans `column_formula_detector.py`, ajouter à `FORMULA_PATTERNS` :

```python
FORMULA_PATTERNS = [
    # ... patterns existants ...
    (r'=\s*([A-Za-z]+)\s*-\s*([A-Za-z]+)', 'subtraction'),
    
    # NOUVEAU PATTERN:
    (r'=\s*SOMME\(([A-Za-z]+):([A-Za-z]+)\)', 'sum_range'),  # =SOMME(A:C)
]
```

Puis ajouter la logique dans `_parse_formula_from_text()`.

### Ajuster le seuil de similarité

Dans `_resolve_column_reference()`, ligne 311 :

```python
if similarity > 0.6:  # Actuellement 60%
    # Modifier à 0.7 pour 70%, etc.
```

---

## 📞 Support

Tout fonctionne automatiquement. Si problème :

1. Vérifier les logs (niveau INFO montre détection, DEBUG montre détails)
2. Exécuter `test_column_formulas.py` pour voir formules détectées
3. Exécuter `test_complete_with_formulas.py` pour test complet
4. Vérifier le fichier Excel généré pour confirmer présence formules

---

## 🎉 Résumé

✅ **Système de détection de formules dans en-têtes : IMPLÉMENTÉ**
✅ **Intégration au remplissage sémantique : COMPLÈTE**
✅ **5 types de formules supportés**
✅ **Résolution intelligente des références**
✅ **Tests créés et prêts**
✅ **100% automatique, aucune configuration nécessaire**

**Le pipeline DSF génère maintenant des formules Excel dynamiques automatiquement !** 🚀
