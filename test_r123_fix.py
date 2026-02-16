"""Test rapide de génération DSF avec les corrections R1/R2/R3"""
import sys
sys.path.insert(0, 'src')

from pathlib import Path
from dsf_pipeline import DSFPipeline, DSFPipelineConfig

# Configuration
config = DSFPipelineConfig(
    template_dsf=Path("templates/DSF Normal standard.xlsx"),
    balance_input=Path("data/BALANCE GULFCAM FINAL 2024.xlsx"),
    dsf_output=Path("output/DSF_TEST_R123_FIXED.xlsx"),
    pfilled_template_path=Path("output/DSF_TEST_R123_prefilled.xlsx"),
    report_json_path=Path("output/reports/test_r123_fixed.json"),
    report_html_path=Path("output/reports/test_r123_fixed.html"),
    apply_calculations=True,
    use_calculation_formulas=True,
    use_smart_general_filler=True
)

print("="*80)
print("🧪 TEST: Génération DSF avec corrections R1/R2/R3")
print("="*80)
print(f"\n📄 Template: {config.template_dsf.name}")
print(f"📊 Balance:  {config.balance_input.name}")
print(f"💾 Output:   {config.dsf_output.name}\n")

# Vérifier les fichiers
missing = []
if not config.template_dsf.exists():
    missing.append(f"Template: {config.template_dsf}")
if not config.balance_input.exists():
    missing.append(f"Balance: {config.balance_input}")

if missing:
    print("❌ Fichiers manquants:")
    for m in missing:
        print(f"   - {m}")
    exit(1)

# Exécuter le pipeline
try:
    print("🚀 Démarrage du pipeline...\n")
    pipeline = DSFPipeline(config)
    artifacts = pipeline.run()
    
    print(f"\n{'='*80}")
    print("✅ DSF GÉNÉRÉ AVEC SUCCÈS")
    print(f"{'='*80}")
    print(f"\n📁 Fichiers créés:")
    print(f"   • DSF final:    {artifacts.dsf_output.name}")
    print(f"   • Prefilled:    {artifacts.prefilled_template.name}")
    if artifacts.report_json_path:
        print(f"   • Rapport JSON: {artifacts.report_json_path.name}")
    if artifacts.report_html_path:
        print(f"   • Rapport HTML: {artifacts.report_html_path.name}")
    
    print(f"\n🔍 Pour vérifier les corrections R1/R2/R3:")
    print(f"   1. Ouvrez: {artifacts.dsf_output}")
    print(f"   2. Consultez les fiches R1, R2, R3")
    print(f"   3. Vérifiez que les données sont correctement placées\n")
    
except Exception as e:
    print(f"\n❌ ERREUR: {e}")
    import traceback
    traceback.print_exc()
    exit(1)
