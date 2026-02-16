# 🧮 COMMENT LE SYSTÈME GÈRE LES COLONNES CALCULÉES

## 📊 Statut Actuel

### ✅ CE QUI EST DÉTECTÉ
Le mapper V7 détecte déjà les colonnes suivantes:

```python
'CALC_TOTAL_MOVEMENTS'   # Colonnes "TOTAL ACQUISITIONS / APPORTS"
'CALC_TOTAL_OPENING'     # Colonnes "TOTAL OUVERTURE / BRUT"
'CALC_RESULT'            # Colonnes "RESULTAT / SOLDE"
```

### ❌ CE QUI MANQUE ACTUELLEMENT
Les calculs ne sont **pas encore appliqués** - les colonnes calculées reçoivent simplement des valeurs de la balance comme les autres colonnes.

---

## 🎯 SOLUTION COMPLÈTE CRÉÉE

### 📦 Nouveau Module: `calculation_manager.py`

**Capacités:**
1. ✅ **Détection lignes TOTAL** - Identifie automatiquement les lignes nécessitant une somme
2. ✅ **Calculs de SOMMES** - `=SUM(D15:D20)` pour les totaux
3. ✅ **Calculs de DIFFÉRENCES** - `Variation = Clôture - Ouverture`
4. ✅ **Calculs de POURCENTAGES** - `Taux = (Nouveau - Ancien) / Ancien * 100`
5. ✅ **Formules Excel** - Génère des formules au lieu de valeurs fixes

---

## 🔧 TYPES DE CALCULS GÉRÉS

### 1️⃣ SOMMES (TOTAL, SOUS-TOTAL)

**Exemple - NOTE 3A:**
```
Ligne 15: Frais de développement       1,000,000    +200,000    1,200,000
Ligne 16: Brevets et licences          3,500,000    +500,000    4,000,000
Ligne 17: Fonds commercial             2,000,000         0      2,000,000
──────────────────────────────────────────────────────────────────────────
Ligne 18: TOTAL IMMOBILISATIONS      [CALCULÉ]    [CALCULÉ]   [CALCULÉ]
          Formule D18: =SUM(D15:D17) → 6,500,000
          Formule E18: =SUM(E15:E17) →   700,000
          Formule F18: =SUM(F15:F17) → 7,200,000
```

**Détection:**
- Libellé contient "TOTAL", "SOUS-TOTAL", "SOMME", "CUMUL"
- Appliqué à toutes colonnes numériques (OPENING, CLOSING, MOVEMENTS)

---

### 2️⃣ DIFFÉRENCES (VARIATION)

**Exemple - Colonne VARIATION:**
```
                          OUVERTURE  CLOTURE   VARIATION
Ligne 20: Actif Total      100,000   120,000   [CALCULÉ]
          Formule G20: =E20-D20 → +20,000
```

**Détection:**
- Header contient "VARIATION", "DIFFÉRENCE", "ÉCART", "RÉSULTAT"
- Calcule: `Colonne_Clôture - Colonne_Ouverture`

---

### 3️⃣ POURCENTAGES (TAUX)

**Exemple - Colonne TAUX VARIATION:**
```
                          OUVERTURE  CLOTURE   TAUX %
Ligne 20: Actif Total      100,000   120,000   [CALCULÉ]
          Formule H20: =(E20-D20)/D20*100 → 20%
```

**Détection:**
- Header contient "%", "POURCENTAGE", "TAUX"
- Calcule: `(Nouveau - Ancien) / Ancien * 100`

---

### 4️⃣ RATIOS

**Exemple - Ratio d'endettement:**
```
Ligne 25: Dettes Totales              50,000
Ligne 30: Actif Total                100,000
Ligne 35: Ratio Endettement          [CALCULÉ]
          Formule D35: =D25/D30 → 0.5 (50%)
```

**Détection:**
- Header contient "RATIO", "COEFFICIENT"
- Configuration manuelle des lignes à diviser

---

## 🚀 INTÉGRATION DANS LE MAPPER

### Étapes pour activer:

```python
# 1. Importer le module
from calculation_manager import get_calculation_manager

# 2. Dans process_sheet_v7(), créer le gestionnaire
calc_mgr = get_calculation_manager(ws, column_info)

# 3. Pour chaque cellule, vérifier si calcul requis
should_calc, calc_type = calc_mgr.should_calculate_cell(row_idx, col_letter, col_type)

if should_calc:
    # Calculer au lieu de remplir depuis balance
    value = calc_mgr.apply_calculation(row_idx, col_letter, col_type, calc_type)
    cell.value = value
else:
    # Remplir depuis balance (code existant)
    ...
```

---

## ⚙️ OPTIONS DE CONFIGURATION

| Option | Valeur | Description |
|--------|--------|-------------|
| **use_excel_formulas** | `True` | Créer formules Excel au lieu de valeurs fixes |
| **calculate_totals_automatically** | `True` | Détecter et calculer les TOTAL automatiquement |
| **preserve_template_formulas** | `True` | Ne jamais écraser formules existantes |
| **min_rows_for_sum** | `2` | Minimum de lignes avant de créer une somme |

---

## 🎓 EXEMPLE COMPLET

### AVANT (Mapper V7 actuel):
```
NOTE 3A - Immobilisations Incorporelles

Ligne 15: Frais développement    1,000,000   200,000   1,200,000
Ligne 16: Brevets licences        3,500,000   500,000   4,000,000
Ligne 17: Fonds commercial        2,000,000         0   2,000,000
Ligne 18: TOTAL                     [VIDE]    [VIDE]     [VIDE]  ❌
```

### APRÈS (Avec calculation_manager):
```
NOTE 3A - Immobilisations Incorporelles

Ligne 15: Frais développement    1,000,000   200,000   1,200,000
Ligne 16: Brevets licences        3,500,000   500,000   4,000,000
Ligne 17: Fonds commercial        2,000,000         0   2,000,000
Ligne 18: TOTAL                   6,500,000   700,000   7,200,000  ✅
          Formules: =SUM(D15:D17), =SUM(E15:E17), =SUM(F15:F17)
```

---

## 📋 CHECKLIST D'IMPLÉMENTATION

- [x] Module `calculation_manager.py` créé
- [x] Détection lignes TOTAL fonctionnelle
- [x] Calcul de sommes implémenté
- [x] Génération formules Excel implémentée
- [x] Calcul différences implémenté
- [x] Calcul pourcentages implémenté
- [ ] **Intégration dans mapper_v7_optimized_formulas.py** (À FAIRE)
- [ ] Tests sur sheets réels (NOTE 3A, NOTE 3B, etc.)
- [ ] Validation des formules générées
- [ ] Documentation utilisateur complète

---

## 🔍 POUR TESTER

```bash
# Test du gestionnaire seul
python calculation_manager.py

# Afficher le guide complet
python GUIDE_CALCULS.py

# Une fois intégré au mapper
python mapper_v7_optimized_formulas.py
```

---

## ❓ Questions Fréquentes

**Q: Les formules Excel sont-elles recalculées automatiquement?**  
R: Oui, si `use_excel_formulas=True`, Excel recalcule automatiquement quand les données changent.

**Q: Que se passe-t-il si un TOTAL ne peut pas être calculé?**  
R: La cellule reste vide (pas d'erreur, mais log dans debug_info).

**Q: Les formules du template sont-elles préservées?**  
R: Oui, si `preserve_template_formulas=True` (par défaut).

**Q: Comment gérer les totaux sur plusieurs niveaux (sous-totaux)?**  
R: Le système détecte automatiquement les niveaux en cherchant le dernier TOTAL/vide au-dessus.

---

## 📞 Support

Pour toute question sur l'implémentation des calculs:
1. Consulter: `GUIDE_CALCULS.py`
2. Lire le code: `calculation_manager.py`
3. Tester: `python calculation_manager.py`
