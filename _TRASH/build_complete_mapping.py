#!/usr/bin/env python3
"""
Build comprehensive DSF-SYSCOHADA mapping from analyze_dsf.json
Maps SYSCOHADA account numbers to DSF cell references
"""
import json
import re
from pathlib import Path
from dataclasses import dataclass
from syscohada_db import get_syscohada_accounts, SYSCOHADA_INDEX

@dataclass
class DSFLineItem:
    """Represents a DSF line item with its reference and cell location"""
    ref_code: str  # e.g., 'CA', 'CB', 'AD', etc.
    label: str
    sheet: str
    brut_row: int = None      # Row for BRUT column (if exists)
    amort_row: int = None     # Row for AMORTIZATION column (if exists)
    net_row: int = None       # Row for NET column (always column D/K)
    note_col: int = None
    description: str = ""

def extract_dsf_structure():
    """Extract DSF structure from dsf_structure.json"""
    with open('dsf_structure.json', 'r', encoding='utf-8') as f:
        dsf_data = json.load(f)
    
    dsf_items = {}  # ref_code -> DSFLineItem
    
    # Process each sheet
    for sheet_name, cells in dsf_data.items():
        print(f"\n=== Processing {sheet_name} ===")
        
        # Extract all cells with REF codes
        ref_pattern = re.compile(r'^([A-Z]{1,2})\d+$')  # Match cell references like A12, CA15, etc.
        
        for cell_ref, cell_value in cells.items():
            # Extract row and column from cell ref (e.g., A12 -> col=A, row=12)
            match = re.match(r'([A-Z]+)(\d+)', cell_ref)
            if not match:
                continue
            
            col_letter = match.group(1)
            row_num = int(match.group(2))
            
            # Column mapping: A/H = REF, B/I = LABEL, D/K = NET values
            is_actif = col_letter in ['A', 'B', 'C', 'D', 'E', 'F', 'G']
            is_passif = col_letter in ['H', 'I', 'J', 'K', 'L']
            
            # Look for REF codes (usually in column A or H, contain uppercase letters + numbers/spaces)
            if (col_letter in ['A', 'H']) and cell_value and len(str(cell_value).strip()) <= 4:
                ref_code = str(cell_value).strip()
                if ref_code and not ref_code.isdigit():  # REF codes like CA, CB, AD, AE
                    # Look ahead for label in next column
                    label_cell = None
                    if col_letter == 'A':
                        label_cell = f'B{row_num}'
                    elif col_letter == 'H':
                        label_cell = f'I{row_num}'
                    
                    label = cells.get(label_cell, '')
                    
                    if ref_code not in dsf_items:
                        dsf_items[ref_code] = DSFLineItem(
                            ref_code=ref_code,
                            label=str(label)[:80],
                            sheet=sheet_name,
                            net_row=row_num
                        )
                        print(f"  REF {ref_code}: {label} (Row {row_num})")
    
    return dsf_items

def build_syscohada_to_dsf_mapping(dsf_items):
    """Build mapping from SYSCOHADA accounts to DSF items based on keywords matching"""
    
    mapping_rules = {
        # Class 1: Capital and resources (100-199)
        '101': ('CA', 'Capital account'),  # Capital
        '102': ('CB', 'Capital non appelé'),  # Non-called capital
        '104': ('CD', 'Primes liées au capital'),  # Capital premiums
        '105': ('CE', 'Écarts de réévaluation'),  # Revaluation differences
        '10': ('CA', 'Capital and reserves'),  # General class 1
        
        # Class 1: Reserves
        '11': ('CF', 'Réserves indisponibles'),
        '110': ('CF', 'Réserves indisponibles'),
        '12': ('CG', 'Réserves libres'),
        '120': ('CG', 'Réserves libres'),
        
        # Class 1: Carried forward earnings
        '13': ('CH', 'Report à nouveau'),
        '130': ('CH', 'Report à nouveau'),
        
        # Class 1: Net result
        '14': ('CJ', 'Résultat net de l\'exercice'),
        '140': ('CJ', 'Résultat net exercice'),
        '141': ('CJ', 'Résultat net exercice'),
        
        # Class 1: Investment subsidies
        '15': ('CL', 'Subventions d\'investissement'),
        '150': ('CL', 'Subventions d\'investissement'),
        
        # Class 1: Regulated provisions
        '16': ('CM', 'Provisions réglementées'),
        '160': ('CM', 'Provisions réglementées'),
        
        # Class 2: Fixed Assets
        '20': ('AD', 'Immobilisations incorporelles'),
        '201': ('AE', 'Frais de développement'),
        '202': ('AF', 'Brevets et logiciels'),
        '203': ('AG', 'Fonds commercial'),
        '204': ('AH', 'Autres immobilisations incorporelles'),
        
        '21': ('AI', 'Immobilisations corporelles'),
        '210': ('AJ', 'Terrains'),
        '211': ('AK', 'Bâtiments'),
        '212': ('AL', 'Aménagements et installations'),
        '213': ('AM', 'Matériel et mobilier'),
        '214': ('AN', 'Matériel de transport'),
        
        '23': ('AP', 'Avances sur immobilisations'),
        '24': ('AQ', 'Immobilisations financières'),
        '241': ('AR', 'Titres de participation'),
        '248': ('AS', 'Autres immobilisations financières'),
        
        # Class 3: Inventory and WIP
        '30': ('BA', 'Matières premières'),
        '31': ('BB', 'Matières consommables'),
        '32': ('BC', 'En-cours de production'),
        '33': ('BD', 'Produits finis'),
        '34': ('BE', 'Marchandises'),
        '35': ('BF', 'Stocks en transit'),
        
        # Class 4: Third parties (Receivables/Payables)
        '40': ('DA', 'Clients et comptes rattachés'),
        '401': ('DA', 'Clients'),
        '408': ('DA', 'Clients - Factures à établir'),
        '411': ('DA', 'Clients - Retenues de garantie'),
        '417': ('DA', 'Clients - Produits non encore facturés'),
        
        '411': ('DA', 'Fournisseurs'),
        '401': ('DA', 'Fournisseurs - Factures reçues'),
        '407': ('DB', 'Fournisseurs - Factures à recevoir'),
        
        # Class 5: Treasury/Financial
        '50': ('BC', 'Disponibilités'),
        '501': ('BC', 'Caisse'),
        '512': ('BC', 'Comptes bancaires'),
        '521': ('BC', 'Comptes courants d\'associés'),
        
        # Class 6: Operating charges
        '60': ('ED', 'Achats'),
        '61': ('EE', 'Services extérieurs'),
        '62': ('EF', 'Autres services extérieurs'),
        '63': ('EG', 'Impôts et taxes'),
        '64': ('EH', 'Charges de personnel'),
        '65': ('EI', 'Autres charges opérationnelles'),
        
        # Class 7: Operating revenues
        '70': ('RD', 'Ventes'),
        '71': ('RE', 'Prestations de services'),
        '72': ('RF', 'Revenus des immeubles'),
        '73': ('RG', 'Autres revenus opérationnels'),
        
        # Debts and financing
        '16': ('DA', 'Emprunts et dettes financières'),
        '161': ('DA', 'Emprunts'),
        '164': ('DB', 'Dettes de location-acquisition'),
        
        # Provisions
        '19': ('DD', 'Provisions pour risques et charges'),
        '191': ('DD', 'Provisions pour litiges'),
        '192': ('DD', 'Provisions pour amendes'),
        '195': ('DD', 'Provisions pour restructuration'),
    }
    
    # Get all SYSCOHADA accounts
    syscohada_accounts = get_syscohada_accounts()
    
    # Build full mapping with multiple strategies
    complete_mapping = {}
    
    for account in syscohada_accounts:
        account_num = account.numero
        
        # Strategy 1: Direct exact match
        if account_num in mapping_rules:
            dsf_ref, desc = mapping_rules[account_num]
            complete_mapping[account_num] = (dsf_ref, desc, 'D')  # Column D for ACTIF net values
        
        # Strategy 2: Class-level match (e.g., '101' matches '10')
        elif account_num[:2] in mapping_rules:
            dsf_ref, desc = mapping_rules[account_num[:2]]
            complete_mapping[account_num] = (dsf_ref, desc, 'D')
        
        # Strategy 3: Prefix match
        else:
            prefix = account_num[:1] or account_num[:2]
            if prefix in mapping_rules:
                dsf_ref, desc = mapping_rules[prefix]
                complete_mapping[account_num] = (dsf_ref, desc, 'D')
    
    return complete_mapping

def create_balance_lignemapping_code(mapping_dict):
    """Generate Python code to create BalanceLigneMapping entries"""
    
    code_lines = [
        "# Generated DSF_MAPPING - Complete SYSCOHADA accounts to DSF mapping",
        "# This maps all SYSCOHADA account numbers to their corresponding DSF cells",
        "",
        "DSF_MAPPING = {",
    ]
    
    # Group by DSF reference for organization
    by_dsf_ref = {}
    for account_num, (dsf_ref, description, col) in mapping_dict.items():
        if dsf_ref not in by_dsf_ref:
            by_dsf_ref[dsf_ref] = []
        by_dsf_ref[dsf_ref].append((account_num, description))
    
    # Generate code with organization
    for dsf_ref in sorted(by_dsf_ref.keys()):
        code_lines.append(f"    # DSF Reference: {dsf_ref}")
        for account_num, description in sorted(by_dsf_ref[dsf_ref]):
            code_lines.append(
                f'    "{account_num}": BalanceLigneMapping('
                f'compte_syscohada="{account_num}", '
                f'dsf_sheet="BILAN PAYSAGE", '
                f'dsf_cell="{dsf_ref}", '
                f'dsf_field="NET", '
                f'is_debit=True, '
                f'multiply_by=1.0, '
                f'condition=None),'
            )
        code_lines.append("")
    
    code_lines.append("}")
    
    return '\n'.join(code_lines)

def main():
    """Main execution"""
    print("Building complete DSF-SYSCOHADA mapping...")
    
    # Extract DSF structure
    dsf_items = extract_dsf_structure()
    print(f"\nFound {len(dsf_items)} DSF line items")
    
    # Build mapping
    print("\nBuilding SYSCOHADA to DSF mapping...")
    mapping = build_syscohada_to_dsf_mapping(dsf_items)
    print(f"Created {len(mapping)} account mappings")
    
    # Generate code
    mapping_code = create_balance_lignemapping_code(mapping)
    
    # Save to file
    output_path = Path('dsf_mapping_generated.py')
    output_path.write_text(mapping_code, encoding='utf-8')
    print(f"\nGenerated mapping code saved to {output_path}")
    
    # Display sample
    print("\n=== Sample Generated Mappings ===")
    lines = mapping_code.split('\n')[10:30]
    print('\n'.join(lines))

if __name__ == '__main__':
    main()
