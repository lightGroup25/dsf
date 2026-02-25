#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Test intégration : Vérifier que les modules fast filler s'importent correctement
et que l'intégration dans le pipeline est en place
"""

import sys
from pathlib import Path

sys.path.insert(0, 'src')

def test_imports():
    """Vérifie que tous les modules peuvent être importés"""
    print("\n[TEST] Vérification des imports...")
    
    try:
        from syscohada_mapper import SyscohadadMapper, Note20Section, Note34Indicator
        print("  [OK] syscohada_mapper imported")
    except ImportError as e:
        print(f"  [FAIL] syscohada_mapper: {e}")
        return False
    
    try:
        from note20_34_fast_filler import Note20FastFiller, Note34FastFiller
        print("  [OK] note20_34_fast_filler imported")
    except ImportError as e:
        print(f"  [FAIL] note20_34_fast_filler: {e}")
        return False
    
    try:
        from fast_filler_integration import apply_fast_fillers_pre_semantic
        print("  [OK] fast_filler_integration imported")
    except ImportError as e:
        print(f"  [FAIL] fast_filler_integration: {e}")
        return False
    
    return True

def test_pipeline_modified():
    """Vérifie que le pipeline DSF a été modifié avec le code fast filler"""
    print("\n[TEST] Vérification des modifications pipeline...")
    
    pipeline_file = Path("src/dsf_pipeline.py")
    if not pipeline_file.exists():
        print(f"  [FAIL] Pipeline file not found: {pipeline_file}")
        return False
    
    content = pipeline_file.read_text()
    
    if "FAST FILLERS PRE-PROCESSING" in content:
        print("  [OK] Fast fillers injection found in pipeline")
    else:
        print("  [FAIL] Fast fillers injection NOT found in pipeline")
        return False
    
    if "apply_fast_fillers_pre_semantic" in content:
        print("  [OK] apply_fast_fillers_pre_semantic called in pipeline")
    else:
        print("  [FAIL] apply_fast_fillers_pre_semantic NOT called in pipeline")
        return False
    
    return True

def test_mapper_functionality():
    """Teste les fonctionnalités du mapper"""
    print("\n[TEST] Fonctionnalités du mapper...")
    
    from syscohada_mapper import SyscohadadMapper
    
    mapper = SyscohadadMapper()
    print("  [OK] Mapper instantiated")
    
    # Teste les colonnes standard
    assert mapper.debit_col_n == 24, "debit_col_n should be 24"
    assert mapper.credit_col_n == 27, "credit_col_n should be 27"
    print("  [OK] Standard columns configured correctly")
    
    return True

def test_fast_fillers_stubs():
    """Teste que les fast fillers peuvent être instantiés"""
    print("\n[TEST] Instanciation des fast fillers...")
    
    from openpyxl import Workbook
    from note20_34_fast_filler import Note20FastFiller, Note34FastFiller
    
    wb = Workbook()
    
    try:
        filler20 = Note20FastFiller(wb)
        print("  [OK] Note20FastFiller instantiated")
    except Exception as e:
        print(f"  [FAIL] Note20FastFiller: {e}")
        return False
    
    try:
        filler34 = Note34FastFiller(wb)
        print("  [OK] Note34FastFiller instantiated")
    except Exception as e:
        print(f"  [FAIL] Note34FastFiller: {e}")
        return False
    
    return True

def main():
    print("\n" + "="*80)
    print("TEST INTEGRATION : Double Mapping SYSCOHADA")
    print("="*80)
    
    all_passed = True
    
    all_passed &= test_imports()
    all_passed &= test_pipeline_modified()
    all_passed &= test_mapper_functionality()
    all_passed &= test_fast_fillers_stubs()
    
    print("\n" + "="*80)
    if all_passed:
        print("STATUS: [OK] All tests passed!")
        print("\nNext step: Run the full pipeline with:")
        print("  python -u src/dsf_pipeline.py")
        print("\nWatch for logs:")
        print("  [FAST FILLERS] Pre-processing NOTE 20 & 34...")
        print("  [FAST FILLERS] Completed in X.XXs")
        return 0
    else:
        print("STATUS: [FAIL] Some tests failed")
        print("\nDebug:")
        print("  - Check that all modules are in src/")
        print("  - Check that dsf_pipeline.py was modified correctly")
        print("  - Check Python imports")
        return 1

if __name__ == "__main__":
    exit(main())
