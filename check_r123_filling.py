"""Vérifie si le semantic_balance_filler remplit des cellules dans R1, R2, R3"""
import json
from pathlib import Path

# Trouver le rapport semantic
reports_dir = Path("output/reports")
if not reports_dir.exists():
    print("❌ Dossier output/reports introuvable")
    exit(1)

# Chercher spécifiquement un rapport semantic ou dsf_report
semantic_reports = list(reports_dir.glob("*semantic*.json")) + list(reports_dir.glob("dsf_report.json"))
if not semantic_reports:
    print("❌ Aucun rapport semantic trouvé")
    exit(1)

latest_report = max(semantic_reports, key=lambda p: p.stat().st_mtime)
print(f"📄 Analyse du rapport: {latest_report.name}\n")

with open(latest_report, 'r', encoding='utf-8', errors='ignore') as f:
    report = json.load(f)

# Chercher les assignments dans R1, R2, R3
r_sheets = ["Fiche R1", "R1", "Fiche R2", "R2", "Fiche R3", "R3"]

# Vérifier si le rapport a des assignments
if "assignments" in report:
    assignments = report["assignments"]
    
    r_assignments = [a for a in assignments if a.get("sheet") in r_sheets]
    
    if r_assignments:
        print("⚠️  PROBLÈME DÉTECTÉ!")
        print(f"Le semantic_balance_filler a rempli {len(r_assignments)} cellules dans R1/R2/R3\n")
        print("=" * 80)
        
        for a in r_assignments[:20]:  # Max 20 exemples
            sheet = a.get("sheet", "?")
            cell = a.get("cell", "?")
            row_label = a.get("row_label", "")
            col_label = a.get("col_label", "")
            amount = a.get("total_amount", 0)
            accounts = a.get("source_accounts",[])
            
            print(f"📍 {sheet:12} {cell:6} | Row: {row_label[:30]:30} | Col: {col_label[:20]:20}")
            print(f"   → Montant: {amount:>15,.0f} | Comptes: {', '.join(accounts[:3])}")
            print()
    else:
        print("✅ Aucune cellule de R1/R2/R3 remplie par semantic_balance_filler")
        print("   (C'est le comportement attendu!)")
else:
    print("⚠️  Structure du rapport non reconnue")
    print(f"   Clés trouvées: {list(report.keys())}")

# Afficher aussi les stats générales
if "validation" in report:
    stats = report["validation"]
    print(f"\n📊 Statistiques générales:")
    print(f"   Total assignments: {stats.get('successful_assignments', 0)}")
    print(f"   Success rate: {stats.get('success_rate', 0):.1f}%")
