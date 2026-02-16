# -*- coding: utf-8 -*-
"""
Documentation: SmartGeneralFiller - Système de reconnaissance automatique
==========================================================================

## Vue d'ensemble

Le `SmartGeneralFiller` est un système intelligent qui remplit automatiquement
les sections générales du DSF (ENTÊTE, R1, R2, R3, NOTE13) en **détectant
automatiquement** les zones de saisie, sans avoir besoin de mapping manuel.

## Fonctionnalités

### 1. Détection Automatique des Zones de Saisie

Le système scanne les feuilles cibles et identifie automatiquement les cellules
à remplir en utilisant plusieurs stratégies :

#### Stratégie A : Pattern "Label:"
```
| Dénomination sociale : | [cellule vide] |
```
→ Détecte la cellule vide à droite du label

#### Stratégie B : Label en gras + cellule en dessous
```
| DÉNOMINATION SOCIALE | ← En gras
| [cellule vide]       | ← Zone de saisie
```

#### Stratégie C : Cellule fusionnée avec zone vide
```
| Dénomination sociale : _________________ |
```
→ Détecte la zone de saisie dans la même cellule fusionnée

### 2. Reconnaissance des Types de Données

Le système infère automatiquement le type de donnée attendu :

- **date** : "Date", "Exercice", "Clos le", "Début", "Fin"
- **currency** : "Capital", "Montant", "Valeur"
- **number** : "Durée", "Mois", "Nombre", "Quantité"
- **email** : "Email", "E-mail"
- **phone** : "Téléphone", "Tél", "Fax"
- **text** : Tous les autres cas

### 3. Matching Intelligent avec DSF_InfosGenerales

Le système utilise deux méthodes pour matcher un label avec un champ :

#### A. Mapping Manuel Prioritaire (100% confiance)
```python
"dénomination sociale" → denomination_sociale
"n° identification"    → num_identification_fiscale
"capital social"       → capital_social
```

#### B. Fuzzy Matching (pour les cas non couverts)
```python
"adresse entreprise" → adresse_complete (score: 0.75)
"tel"                → telephone (score: 0.68)
```

Seuil de confiance minimum : **50%**

### 4. Format Automatique

Le système formate automatiquement les valeurs selon le type détecté :

- **Dates** : `01/01/2024` → `"01-01-2024"`
- **Currency** : `1500000` → `1500000.0`
- **Text** : Préservé tel quel
- **Cellules fusionnées** : `"Label : Valeur"`

## Utilisation

### Dans le Pipeline

```python
from dsf_pipeline import DSFPipeline, DSFPipelineConfig

config = DSFPipelineConfig(
    template_dsf=Path("templates/DSF Normal standard.xlsx"),
    balance_input=Path("input/balance.xlsx"),
    use_smart_general_filler=True,  # Active le système intelligent
)

pipeline = DSFPipeline(config)
artifacts = pipeline.run()
```

### Standalone

```python
from smart_general_filler import SmartGeneralFiller
from dsf_general_info import get_gulfcam_config

# Charger les infos de l'entreprise
info = get_gulfcam_config()

# Créer le filler
filler = SmartGeneralFiller("templates/DSF Normal standard.xlsx")

# 1. Détecter les zones
filler.load()
print(f"{len(filler.input_zones)} zones détectées")

# 2. Remplir automatiquement
filled = filler.fill(info)
print(f"{filled} zones remplies")

# 3. Sauvegarder
filler.save("output/dsf_filled.xlsx")

# 4. Rapport détaillé
report = filler.get_report()
print(f"Taux de remplissage: {report['fill_rate']*100:.1f}%")

filler.close()
```

## Configuration dans le Pipeline

### Activer le système intelligent (par défaut)
```python
config = DSFPipelineConfig(
    use_smart_general_filler=True  # Reconnaissance automatique
)
```

### Utiliser l'ancien système (mapping manuel)
```python
config = DSFPipelineConfig(
    use_smart_general_filler=False  # Utilise dsf_prefill_mapping.json
)
```

## Rapport de Remplissage

Le système génère un rapport détaillé :

```python
report = filler.get_report()
# {
#     "total_zones": 45,
#     "total_filled": 38,
#     "fill_rate": 0.84,
#     "by_sheet": {
#         "ENTETE": {
#             "detected": 12,
#             "filled": 10,
#             "zones": [
#                 {
#                     "cell": "A2",
#                     "label": "Dénomination sociale",
#                     "type": "text",
#                     "matched_field": "denomination_sociale",
#                     "confidence": 1.0
#                 },
#                 ...
#             ]
#         },
#         ...
#     }
# }
```

## Avantages vs Ancien Système

| Critère | Ancien (DSFGeneralPrefiller) | Nouveau (SmartGeneralFiller) |
|---------|------------------------------|------------------------------|
| **Configuration** | Mapping JSON manuel (226 lignes) | Aucune configuration nécessaire |
| **Flexibilité** | Rigide, cellules fixes | S'adapte à la structure du template |
| **Maintenance** | Mise à jour manuelle du JSON | Automatique |
| **Nouveaux champs** | Ajouter manuellement au JSON | Détecté automatiquement |
| **Template modifié** | Mapper à nouveau | Continue de fonctionner |
| **Taux de couverture** | Dépend du mapping | Détecte toutes les zones reconnaissables |

## Test du Système

```bash
# Test standalone
python test_smart_filler.py

# Test via pipeline complet
python src/dsf_pipeline.py

# Test via interface graphique
python src/dsf_desktop_ui.py
```

## Patterns de Labels Reconnus

Le système reconnaît automatiquement plus de 20 patterns :

```
✓ dénomination sociale
✓ raison sociale  
✓ sigle
✓ adresse
✓ identification fiscale
✓ numéro / n° / nif / niu
✓ téléphone / telephone / tél
✓ fax
✓ email / e-mail
✓ exercice
✓ capital
✓ date
✓ durée / duree
✓ activité / activite
✓ forme juridique
✓ système comptable / systeme
✓ centre dépôt / depot
✓ ville
✓ pays
✓ secteur
✓ code
```

## Logs et Débogage

Le système génère des logs détaillés :

```
INFO - Loading template: templates/DSF Normal standard.xlsx
INFO - Scanning sheet: ENTETE
INFO -   → 12 input zones detected
INFO - Scanning sheet: Fiche R1
INFO -   → 8 input zones detected
INFO - Total input zones detected: 45
DEBUG - ✓ Filled ENTETE!A2: denomination_sociale = GULFCAM SA (conf=1.00)
DEBUG - ✓ Filled ENTETE!B5: num_identification_fiscale = M061200000001 (conf=1.00)
INFO - Filled 38/45 zones
```

## Structure de InputZone

```python
class InputZone:
    sheet: str              # Nom de la feuille
    cell_ref: str          # Référence de cellule (ex: "A2")
    row: int               # Numéro de ligne
    col: int               # Numéro de colonne
    label: str             # Label extrait (ex: "Dénomination sociale")
    label_cell: str        # Cellule du label
    zone_type: str         # Type: text, date, number, currency, email, phone
    context: str           # Contexte complet de la cellule
    matched_field: str     # Champ associé (ex: "denomination_sociale")
    confidence: float      # Score de confiance (0.0 à 1.0)
```

## Limitations et Améliorations Futures

### Limitations Actuelles
- Scan limité aux 50 premières lignes et 10 premières colonnes
- Seuil de confiance fixe (50%)
- Pas de gestion des tableaux complexes (dirigeants, actionnaires)

### Améliorations Prévues
- Support des tableaux multi-lignes
- Apprentissage par feedback utilisateur
- Export des mappings détectés pour validation
- Mode interactif pour validation manuelle

## Intégration

Le système est **100% intégré** dans le pipeline :

```
1. Template chargé
   ↓
2. ✓ SmartGeneralFiller détecte zones (ENTÊTE/R1/R2/R3/NOTE13)
   ↓
3. ✓ Matching automatique avec DSF_InfosGenerales
   ↓
4. ✓ Remplissage avec formats corrects
   ↓
5. Remplissage sémantique balance (reste du DSF)
   ↓
6. Application calculs automatiques (formules)
   ↓
7. Génération rapports
```

## Questions Fréquentes

**Q: Que se passe-t-il si une zone n'est pas détectée ?**  
R: Le système log un warning. Vous pouvez ajouter le pattern au mapping manuel.

**Q: Puis-je forcer un mapping spécifique ?**  
R: Oui, ajoutez-le dans `manual_mappings` de `_match_field()`.

**Q: Le système fonctionne-t-il avec d'autres templates ?**  
R: Oui, tant que les labels suivent des patterns reconnus.

**Q: Comment désactiver le système ?**  
R: `use_smart_general_filler=False` dans la config.

**Q: Les performances sont-elles bonnes ?**  
R: Oui, détection + remplissage < 2 secondes pour un template standard.

## Fichiers Créés

- `src/smart_general_filler.py` - Module principal (600+ lignes)
- `test_smart_filler.py` - Script de test standalone
- `SMART_FILLER_DOC.md` - Cette documentation

## Maintenance

Aucune maintenance nécessaire pour l'utilisation normale.

Pour ajouter un nouveau pattern de label :
1. Ajouter le regex dans `LABEL_PATTERNS`
2. Ajouter le mapping dans `manual_mappings` si nécessaire
3. Tester avec `test_smart_filler.py`

## Support

Pour tout problème :
1. Vérifier les logs (niveau DEBUG pour détails)
2. Générer un rapport avec `get_report()`
3. Tester en mode standalone avec `test_smart_filler.py`
4. Comparer avec l'ancien système (`use_smart_general_filler=False`)
