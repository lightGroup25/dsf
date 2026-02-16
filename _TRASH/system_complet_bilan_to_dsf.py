"""
SYSTEM COMPLET: Bilan GULFCAM → DSF Rempli
============================================

PROCESSUS:
1. Lire bilan GULFCAM
2. Extraire comptes 2 chiffres (TOTAUX de classe)
3. Mapper au DSF (cellules synthèse)
4. Remplir le DSF template avec dsf_prefiller_production.py

IMPORTANT:
- Comptes 2 chiffres (10, 20, 30, etc.) = totaux classe → DSF synthèse
- Autres comptes = détails → pas directement au DSF (correspondance)
"""

import logging
from openpyxl import load_workbook
from pathlib import Path
from datetime import date
from dsf_prefiller_production import DSFPrefillerProduction, create_gulfcam_company_info

logging.basicConfig(level=logging.INFO, format='%(message)s')
logger = logging.getLogger(__name__)


# =============================================================================
# ÉTAPE 1: EXTRACTION BILAN GULFCAM
# =============================================================================

class BalanceReader:
    """Lit le bilan GULFCAM et extrait les comptes avec TOTAUX 2 chiffres"""
    
    def __init__(self, balance_file: str):
        self.balance_file = Path(balance_file)
        self.wb = None
        self.sheet = None
        self.headers = {}
        self.balance_data = []
        
    def load(self):
        """Charge le fichier bilan"""
        logger.info(f"📂 Lecture bilan: {self.balance_file.name}")
        
        if not self.balance_file.exists():
            raise FileNotFoundError(f"Fichier introuvable: {self.balance_file}")
        
        self.wb = load_workbook(self.balance_file, data_only=True)
        self.sheet = self.wb.active
        
        # Auto-détect colonnes
        self._detect_columns()
        logger.info(f"  ✓ Colonnes détectées: {self.headers}")
        
        return self
    
    def _detect_columns(self):
        """Détecte automatiquement les colonnes Compte, Débit, Crédit"""
        for col in range(1, self.sheet.max_column + 1):
            header = str(self.sheet.cell(row=1, column=col).value or "").lower()
            
            if "compte" in header:
                self.headers["compte"] = col
            elif "debit" in header or "débit" in header:
                self.headers["debit"] = col
            elif "credit" in header or "crédit" in header:
                self.headers["credit"] = col
        
        if len(self.headers) < 3:
            raise ValueError("❌ Impossible de détecter les colonnes (Compte, Débit, Crédit)")
    
    def _safe_float(self, value):
        """Convertit une valeur en float"""
        try:
            if value is None:
                return 0.0
            return float(value)
        except:
            return 0.0
    
    def read_all_accounts(self):
        """Lit tous les comptes du bilan"""
        logger.info("\n📋 Extraction des comptes...")
        
        self.balance_data = []
        
        for row in range(2, self.sheet.max_row + 1):
            compte_cell = self.sheet.cell(row=row, column=self.headers["compte"])
            compte = str(compte_cell.value or "").strip()
            
            if not compte or compte == "None":
                continue
            
            debit = self._safe_float(self.sheet.cell(row=row, column=self.headers["debit"]).value)
            credit = self._safe_float(self.sheet.cell(row=row, column=self.headers["credit"]).value)
            
            solde_debit = debit  # Débit positif
            solde_credit = -credit  # Crédit négatif
            
            self.balance_data.append({
                "compte": compte,
                "debit": debit,
                "credit": credit,
                "solde_debit": solde_debit,
                "solde_credit": solde_credit,
                "solde": solde_debit + solde_credit
            })
        
        logger.info(f"  ✓ {len(self.balance_data)} comptes lus")
        return self
    
    def extract_two_digit_totals(self):
        """
        Extrait les comptes 2 chiffres (totaux de classe)
        
        Exemple:
        - 10 = Total capital
        - 20 = Total immobilisations
        - 30 = Total stocks
        - 40 = Total tiers
        - 50 = Total trésorerie
        - 60 = Total charges
        - 70 = Total produits
        """
        logger.info("\n💾 Extraction des TOTAUX (comptes 2 chiffres)...")
        
        totals = {}
        
        for entry in self.balance_data:
            compte = entry["compte"]
            
            # Comptes 2 chiffres uniquement
            if len(compte) == 2 and compte.isdigit():
                totals[compte] = entry["solde"]
                logger.info(f"  ✓ Compte {compte}: {entry['solde']:>12,.0f}")
        
        logger.info(f"\n  Total: {len(totals)} comptes 2 chiffres trouvés")
        return totals
    
    def get_detailed_accounts(self):
        """Retourne TOUS les comptes (pour vérification)"""
        return [e for e in self.balance_data if len(e["compte"]) > 2]
    
    def close(self):
        """Ferme le fichier"""
        if self.wb:
            self.wb.close()


# =============================================================================
# ÉTAPE 2: MAPPING COMPTES 2 CHIFFRES → DSF (SYNTHÈSE)
# =============================================================================

def create_totals_mapping():
    """
    Crée le mapping comptes 2 chiffres → cellules DSF synthèse
    
    Exemple de structure DSF:
    - Bilan (BILAN PAYSAGE): Actif/Passif
    - Compte Résultat: Charges/Produits
    """
    
    mapping = {
        # ACTIF (Classe 2, 3, 4, 5)
        "10": {"sheet": "BILAN PAYSAGE", "cell": "D5", "type": "equity", "label": "Capital souscrit"},
        "12": {"sheet": "BILAN PAYSAGE", "cell": "D7", "type": "equity", "label": "Primes d'émission"},
        "13": {"sheet": "BILAN PAYSAGE", "cell": "D8", "type": "equity", "label": "Écarts de réévaluation"},
        "14": {"sheet": "BILAN PAYSAGE", "cell": "D9", "type": "equity", "label": "Réserves"},
        "19": {"sheet": "BILAN PAYSAGE", "cell": "D12", "type": "equity", "label": "Résultats reportés"},
        
        # ACTIF IMMOBILISÉ (Classe 2)
        "20": {"sheet": "BILAN PAYSAGE", "cell": "B5", "type": "asset", "label": "Immobilisations incorporelles"},
        "21": {"sheet": "BILAN PAYSAGE", "cell": "B7", "type": "asset", "label": "Immobilisations corporelles"},
        "27": {"sheet": "BILAN PAYSAGE", "cell": "B10", "type": "asset", "label": "Amortissements"},
        
        # STOCKS (Classe 3)
        "30": {"sheet": "BILAN PAYSAGE", "cell": "B13", "type": "asset", "label": "Stocks"},
        "39": {"sheet": "BILAN PAYSAGE", "cell": "B14", "type": "asset", "label": "Provision stocks"},
        
        # TIERS (Classe 4)
        "40": {"sheet": "BILAN PAYSAGE", "cell": "B16", "type": "asset", "label": "Créances clients"},
        "41": {"sheet": "BILAN PAYSAGE", "cell": "B17", "type": "asset", "label": "Autres créances"},
        "42": {"sheet": "BILAN PAYSAGE", "cell": "D15", "type": "liability", "label": "Dettes fournisseurs"},
        "44": {"sheet": "BILAN PAYSAGE", "cell": "D16", "type": "liability", "label": "Impôts/Taxes"},
        "46": {"sheet": "BILAN PAYSAGE", "cell": "D17", "type": "liability", "label": "Dettes sociales"},
        
        # TRÉSORERIE (Classe 5)
        "50": {"sheet": "BILAN PAYSAGE", "cell": "B19", "type": "asset", "label": "Valeurs mobilières"},
        "51": {"sheet": "BILAN PAYSAGE", "cell": "B20", "type": "asset", "label": "Banques"},
        "55": {"sheet": "BILAN PAYSAGE", "cell": "B21", "type": "asset", "label": "Caisse"},
        "56": {"sheet": "BILAN PAYSAGE", "cell": "D20", "type": "liability", "label": "Dettes financières"},
        
        # CHARGES (Classe 6)
        "60": {"sheet": "COMPTE RESULTAT", "cell": "D8", "type": "expense", "label": "Charges d'exploitation"},
        "61": {"sheet": "COMPTE RESULTAT", "cell": "D9", "type": "expense", "label": "Charges externes"},
        "62": {"sheet": "COMPTE RESULTAT", "cell": "D10", "type": "expense", "label": "Charges personnel"},
        "63": {"sheet": "COMPTE RESULTAT", "cell": "D11", "type": "expense", "label": "Impôts/Taxes"},
        "67": {"sheet": "COMPTE RESULTAT", "cell": "D15", "type": "expense", "label": "Charges financières"},
        "69": {"sheet": "COMPTE RESULTAT", "cell": "D17", "type": "expense", "label": "Charges exceptionnelles"},
        
        # PRODUITS (Classe 7)
        "70": {"sheet": "COMPTE RESULTAT", "cell": "C8", "type": "revenue", "label": "Produits exploitation"},
        "71": {"sheet": "COMPTE RESULTAT", "cell": "C9", "type": "revenue", "label": "Produits externes"},
        "72": {"sheet": "COMPTE RESULTAT", "cell": "C10", "type": "revenue", "label": "Variation stocks"},
        "75": {"sheet": "COMPTE RESULTAT", "cell": "C15", "type": "revenue", "label": "Produits financiers"},
        "79": {"sheet": "COMPTE RESULTAT", "cell": "C17", "type": "revenue", "label": "Produits exceptionnels"},
    }
    
    return mapping


# =============================================================================
# ÉTAPE 3: REMPLISSAGE DSF
# =============================================================================

def fill_dsf_from_balance(balance_totals: dict, output_file: str = "DSF_GULFCAM_COMPLETE.xlsx"):
    """
    Remplit le DSF avec les totaux du bilan
    
    Processus:
    1. Crée prefiller
    2. Remplit info générale (ENTETE)
    3. Mappe totaux aux cellules DSF
    4. Sauvegarde
    """
    
    logger.info("\n" + "="*70)
    logger.info("🔄 REMPLISSAGE DSF")
    logger.info("="*70)
    
    # Initialiser prefiller
    prefiller = DSFPrefillerProduction()
    prefiller.load_template()
    prefiller.load_mapping()
    
    # 1️⃣ Remplir infos ENTETE
    logger.info("\n1️⃣  Remplissage ENTETE (infos générales)...")
    company_info = create_gulfcam_company_info()
    prefiller.fill_entete(company_info)
    prefiller.fill_info_generales(company_info)
    logger.info("  ✓ Infos générales remplies")
    
    # 2️⃣ Mapper totaux aux cellules DSF
    logger.info("\n2️⃣  Mapping totaux aux cellules DSF...")
    mapping = create_totals_mapping()
    
    # Grouper par feuille
    sheets_data = {}
    
    for compte, valeur in balance_totals.items():
        if compte not in mapping:
            logger.warning(f"  ⚠ Compte {compte} non mappé")
            continue
        
        map_info = mapping[compte]
        sheet_name = map_info["sheet"]
        cell_ref = map_info["cell"]
        
        if sheet_name not in sheets_data:
            sheets_data[sheet_name] = {}
        
        sheets_data[sheet_name][cell_ref] = valeur
        logger.info(f"  ✓ {compte} ({map_info['label']:30s}) → {sheet_name}!{cell_ref} = {valeur:>12,.0f}")
    
    # 3️⃣ Écrire dans le DSF
    logger.info("\n3️⃣  Écriture dans le DSF...")
    for sheet_name, cells_data in sheets_data.items():
        count = prefiller.fill_sheet_from_mapping(sheet_name, cells_data)
        logger.info(f"  ✓ {sheet_name}: {count} cellules remplies")
    
    # 4️⃣ Sauvegarder
    logger.info(f"\n4️⃣  Sauvegarde...")
    prefiller.save(output_file)
    
    logger.info("\n" + "="*70)
    logger.info(f"✅ DSF REMPLI: {output_file}")
    logger.info("="*70)


# =============================================================================
# MAIN: EXÉCUTION COMPLÈTE
# =============================================================================

def main():
    """Exécution end-to-end complète"""
    
    logger.info("\n" + "="*70)
    logger.info("SYSTÈME COMPLET: Bilan GULFCAM → DSF")
    logger.info("="*70)
    
    # ÉTAPE 1: Lire bilan
    balance_file = "BALANCE AU 31 12 2024 VERSION DU 26 05 25.xlsx"
    
    reader = BalanceReader(balance_file)
    reader.load().read_all_accounts()
    
    # ÉTAPE 2: Extraire totaux 2 chiffres
    totals = reader.extract_two_digit_totals()
    
    # ÉTAPE 3: Vérifier équilibre
    logger.info("\n✓ Vérification équilibre...")
    
    # Actif = (Immob + Stocks + Tiers + Trésorerie)
    actif = totals.get("20", 0) + totals.get("30", 0) + totals.get("40", 0) + totals.get("50", 0)
    
    # Passif = (Capitaux propres + Dettes)
    passif = totals.get("10", 0) + totals.get("42", 0) + totals.get("44", 0) + totals.get("56", 0)
    
    logger.info(f"  Actif:  {actif:>12,.0f}")
    logger.info(f"  Passif: {passif:>12,.0f}")
    logger.info(f"  Écart:  {abs(actif - passif):>12,.0f}")
    
    # ÉTAPE 4: Remplir DSF
    fill_dsf_from_balance(totals)
    
    # ÉTAPE 5: Résumé
    logger.info("\n" + "="*70)
    logger.info("RÉSUMÉ")
    logger.info("="*70)
    logger.info(f"✅ Comptes 2 chiffres: {len(totals)}")
    logger.info(f"✅ Total comptes lus: {len(reader.balance_data)}")
    logger.info(f"✅ Sortie: DSF_GULFCAM_COMPLETE.xlsx")
    logger.info("="*70 + "\n")
    
    reader.close()


if __name__ == "__main__":
    main()
