#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Vérification rapide des imports et de la structure."""

import sys
from pathlib import Path

# Ajouter le répertoire courant au path
sys.path.insert(0, str(Path(__file__).parent))

print("🔍 Vérification des imports...")

try:
    from src.dsf_desktop_ui import (
        LightDSFMapperWindow,
        AppSettings,
        get_colors,
        get_styles,
        DropArea,
        AnimatedStackedWidget,
        ExcelWidget,
        HeaderWidget,
        CompanyInfoCard,
        DocumentsCard,
        SubmissionStatusCard,
        LandingView,
        BalancePreviewView,
        DsfPreviewView,
    )
    print("✅ Tous les imports réussis!")
    print()
    
    # Vérifier AppSettings
    print("📋 Test AppSettings...")
    settings = AppSettings()
    print(f"   • Dark mode: {settings.dark_mode}")
    print(f"   • Sidebar collapsed: {settings.sidebar_collapsed}")
    print(f"   • Window: {settings.window_width}x{settings.window_height}")
    print(f"   • Recent files: {len(settings.recent_files)}")
    print("   ✅ AppSettings OK")
    print()
    
    # Vérifier les couleurs
    print("🎨 Test des couleurs...")
    light_colors = get_colors(False)
    dark_colors = get_colors(True)
    print(f"   • Light primary: {light_colors['primary']}")
    print(f"   • Dark primary: {dark_colors['primary']}")
    print("   ✅ Couleurs OK")
    print()
    
    # Vérifier les styles
    print("💅 Test des styles...")
    light_styles = get_styles(light_colors)
    dark_styles = get_styles(dark_colors)
    print(f"   • Light styles keys: {len(light_styles)}")
    print(f"   • Dark styles keys: {len(dark_styles)}")
    print("   ✅ Styles OK")
    print()
    
    print("✅ TOUS LES TESTS PASSED!")
    print()
    print("🚀 L'application est prête à être lancée!")
    print()
    print("   Commande: python test_ui_pro.py")
    
except Exception as e:
    print(f"❌ ERREUR: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)
