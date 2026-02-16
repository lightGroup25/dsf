#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""LIGHT DSF MAPPER – Design GULFCAM avec Landing Page Animée et Excel Intégré"""

from __future__ import annotations

import json
import os
import shutil
import time
import tempfile
from pathlib import Path
from typing import Dict, List, Optional
from dataclasses import dataclass, asdict
import pythoncom
import win32com.client
from win32com.client import constants
import win32gui
import win32con

from PySide6.QtCore import Qt, QTimer, QSize, QPropertyAnimation, QEasingCurve, QPoint, QRect
from PySide6.QtGui import QFont, QAction, QColor, QIcon, QKeySequence, QShortcut, QPixmap, QPainter
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSizePolicy,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
    QFrame,
    QProgressBar,
    QGraphicsDropShadowEffect,
    QStatusBar,
    QGridLayout,
    QScrollArea,
)

# FontAwesome désactivé
FONTS_AVAILABLE = False
qta = None

# Lazy imports
try:
    from fill_dsf_from_balance import BalanceReader
except ImportError:
    BalanceReader = None

try:
    from dsf_pipeline import DSFPipeline, DSFPipelineConfig, PipelineArtifacts
except ImportError as e:
    print(f"Warning: Could not import DSF modules: {e}")
    DSFPipeline = None
    DSFPipelineConfig = None
    PipelineArtifacts = None

# CHARTE GRAPHIQUE GULFCAM
GULFCAM_COLORS = {
    'primary': '#003366',      # Bleu profond GULFCAM
    'primary_light': '#0047b3', # Bleu plus clair
    'secondary': '#f59e0b',    # Orange/ambre GULFCAM (accent)
    'secondary_light': '#fbbf24', # Orange plus clair
    'accent': '#ff6b35',        # Orange vif
    'success': '#10b981',       # Vert
    'danger': '#ef4444',        # Rouge
    'warning': '#f59e0b',       # Orange
    'light': '#f8fafc',         # Blanc cassé
    'dark': '#0f172a',          # Bleu très foncé
    'gray': '#64748b',          # Gris
    'light-gray': '#e2e8f0',    # Gris clair
    'card-bg': '#ffffff',       # Blanc
    'gradient_start': '#003366', # Dégradé début
    'gradient_end': '#0047b3',   # Dégradé fin
}

# Pour le mode sombre
GULFCAM_DARK_COLORS = {
    'primary': '#0047b3',       # Bleu plus clair pour mode sombre
    'primary_light': '#3b82f6',
    'secondary': '#f59e0b',     # Orange conservé
    'secondary_light': '#fbbf24',
    'accent': '#ff6b35',
    'success': '#10b981',
    'danger': '#ef4444',
    'warning': '#f59e0b',
    'light': '#1f2937',
    'dark': '#ffffff',
    'gray': '#9ca3af',
    'light-gray': '#374151',
    'card-bg': '#1f2937',
    'gradient_start': '#1e3a5f',
    'gradient_end': '#2563eb',
}

# Logo GULFCAM
GULFCAM_LOGO_TEXT = "GULFCAM"
GULFCAM_TAGLINE = "S'investir pour vous"

# App Settings
SETTINGS_FILE = Path.home() / ".light_dsf_mapper_settings.json"

@dataclass
class AppSettings:
    """Paramètres de l'application persistants."""
    dark_mode: bool = False
    window_width: int = 1400
    window_height: int = 900
    recent_files: List[str] = None
    
    def __post_init__(self):
        if self.recent_files is None:
            self.recent_files = []
    
    def save(self):
        try:
            with open(SETTINGS_FILE, 'w') as f:
                json.dump(asdict(self), f, indent=2)
        except Exception as e:
            print(f"Erreur sauvegarde settings: {e}")
    
    @staticmethod
    def load():
        if SETTINGS_FILE.exists():
            try:
                with open(SETTINGS_FILE, 'r') as f:
                    data = json.load(f)
                    # Filtrer les clés inconnues
                    valid_keys = ['dark_mode', 'window_width', 'window_height', 'recent_files']
                    filtered_data = {k: v for k, v in data.items() if k in valid_keys}
                    return AppSettings(**filtered_data)
            except Exception as e:
                print(f"Erreur chargement settings: {e}")
        return AppSettings()

def get_gulfcam_colors(dark_mode=False):
    """Retourne les couleurs GULFCAM selon le thème"""
    return GULFCAM_DARK_COLORS if dark_mode else GULFCAM_COLORS

class SplashScreen(QWidget):
    """Landing page animée au démarrage avec branding GULFCAM"""
    
    def __init__(self, callback):
        super().__init__()
        self.callback = callback
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        
        # Taille de l'écran
        screen = QApplication.primaryScreen().geometry()
        self.setGeometry(screen)
        
        # Animation d'opacité
        self.opacity_animation = QPropertyAnimation(self, b"windowOpacity")
        self.opacity_animation.setDuration(800)
        self.opacity_animation.setStartValue(0.0)
        self.opacity_animation.setEndValue(1.0)
        self.opacity_animation.setEasingCurve(QEasingCurve.OutCubic)
        
        # Configuration UI
        self.setup_ui()
        
        # Timer pour fermer le splash
        QTimer.singleShot(5000, self.close_splash)
        
    def setup_ui(self):
        """Configure l'interface du splash screen avec branding GULFCAM"""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setAlignment(Qt.AlignCenter)
        
        # Conteneur principal avec fond dégradé GULFCAM
        container = QFrame()
        container.setStyleSheet(f"""
            QFrame {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                    stop:0 {GULFCAM_COLORS['gradient_start']}, 
                    stop:0.7 {GULFCAM_COLORS['gradient_end']});
                border-radius: 0px;
            }}
        """)
        container.setFixedSize(self.width(), self.height())
        
        container_layout = QVBoxLayout(container)
        container_layout.setAlignment(Qt.AlignCenter)
        container_layout.setSpacing(20)
        
        # Logo GULFCAM (version texte stylisé)
        logo_container = QFrame()
        logo_layout = QVBoxLayout(logo_container)
        logo_layout.setAlignment(Qt.AlignCenter)
        logo_layout.setSpacing(8)
        
        # Nom GULFCAM avec style
        gulfcam_label = QLabel(GULFCAM_LOGO_TEXT)
        gulfcam_label.setStyleSheet("""
            QLabel {
                color: white;
                font-size: 64px;
                font-weight: 800;
                letter-spacing: 8px;
                background: transparent;
            }
        """)
        
        # Tagline
        tagline_label = QLabel(GULFCAM_TAGLINE)
        tagline_label.setStyleSheet("""
            QLabel {
                color: rgba(255,255,255,0.9);
                font-size: 18px;
                font-weight: 400;
                letter-spacing: 2px;
                background: transparent;
                font-style: italic;
            }
        """)
        
        logo_layout.addWidget(gulfcam_label)
        logo_layout.addWidget(tagline_label)
        
        # Titre DSF BY LIGHT GROUP
        dsf_label = QLabel("DSF")
        dsf_label.setStyleSheet("""
            QLabel {
                color: white;
                font-size: 48px;
                font-weight: 700;
                letter-spacing: 4px;
                background: transparent;
                margin-top: 20px;
            }
        """)
        
        light_label = QLabel("BY LIGHT GROUP")
        light_label.setStyleSheet("""
            QLabel {
                color: rgba(255,255,255,0.8);
                font-size: 24px;
                font-weight: 300;
                letter-spacing: 4px;
                background: transparent;
            }
        """)
        
        # Barre de progression animée
        self.progress_bar = QFrame()
        self.progress_bar.setFixedSize(300, 4)
        self.progress_bar.setStyleSheet(f"""
            QFrame {{
                background-color: rgba(255,255,255,0.2);
                border-radius: 2px;
            }}
        """)
        
        self.progress_fill = QFrame(self.progress_bar)
        self.progress_fill.setFixedSize(0, 4)
        self.progress_fill.setStyleSheet(f"""
            QFrame {{
                background-color: {GULFCAM_COLORS['secondary']};
                border-radius: 2px;
            }}
        """)
        
        # Animation de la barre de progression
        self.progress_animation = QPropertyAnimation(self.progress_fill, b"geometry")
        self.progress_animation.setDuration(4800)
        self.progress_animation.setStartValue(QRect(0, 0, 0, 4))
        self.progress_animation.setEndValue(QRect(0, 0, 300, 4))
        self.progress_animation.setEasingCurve(QEasingCurve.Linear)
        
        # Texte de chargement
        self.loading_label = QLabel("Chargement de l'application...")
        self.loading_label.setStyleSheet("""
            QLabel {
                color: rgba(255,255,255,0.7);
                font-size: 14px;
                font-weight: 300;
                background: transparent;
            }
        """)
        
        container_layout.addWidget(logo_container)
        container_layout.addSpacing(40)
        container_layout.addWidget(dsf_label)
        container_layout.addWidget(light_label)
        container_layout.addSpacing(60)
        container_layout.addWidget(self.progress_bar, 0, Qt.AlignCenter)
        container_layout.addWidget(self.loading_label)
        
        layout.addWidget(container)
        
        # Démarrer les animations
        self.start_animations()
        
    def start_animations(self):
        """Démarre toutes les animations"""
        self.opacity_animation.start()
        self.progress_animation.start()
        
    def close_splash(self):
        """Ferme le splash screen et appelle le callback"""
        self.close()
        self.callback()

class ModernCard(QFrame):
    """Carte moderne avec effets visuels et branding GULFCAM"""
    
    def __init__(self, parent=None, colors=None):
        super().__init__(parent)
        self.colors = colors or GULFCAM_COLORS
        self.setup_ui()
        
    def setup_ui(self):
        self.setFrameShape(QFrame.NoFrame)
        
        # Ombre portée
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(20)
        shadow.setColor(QColor(0, 51, 102, 30))
        shadow.setOffset(0, 4)
        self.setGraphicsEffect(shadow)
        
        self.update_style()
        
    def update_style(self):
        """Met à jour le style avec les couleurs actuelles"""
        self.setStyleSheet(f"""
            ModernCard {{
                background-color: {self.colors['card-bg']};
                border-radius: 16px;
                border: 1px solid {self.colors['light-gray']};
            }}
        """)
        
    def set_colors(self, colors):
        """Met à jour les couleurs"""
        self.colors = colors
        self.update_style()

class AnimatedButton(QPushButton):
    """Bouton avec animations au survol et couleurs GULFCAM"""
    
    def __init__(self, text="", parent=None, primary=True, colors=None):
        super().__init__(text, parent)
        self.colors = colors or GULFCAM_COLORS
        self.primary = primary
        self.setCursor(Qt.PointingHandCursor)
        self.setMinimumHeight(44)
        self.setMinimumWidth(120)
        self.update_style()
        
    def update_style(self):
        """Met à jour le style selon le type de bouton"""
        if self.primary:
            self.setStyleSheet(f"""
                QPushButton {{
                    background-color: {self.colors['primary']};
                    color: white;
                    border: none;
                    border-radius: 8px;
                    padding: 12px 24px;
                    font-weight: 600;
                    font-size: 14px;
                }}
                QPushButton:hover {{
                    background-color: {self.colors['primary_light']};
                }}
                QPushButton:pressed {{
                    background-color: {self.colors['secondary']};
                }}
                QPushButton:disabled {{
                    background-color: {self.colors['light-gray']};
                    color: {self.colors['gray']};
                }}
            """)
        else:
            self.setStyleSheet(f"""
                QPushButton {{
                    background-color: transparent;
                    color: {self.colors['primary']};
                    border: 2px solid {self.colors['primary']};
                    border-radius: 8px;
                    padding: 12px 24px;
                    font-weight: 600;
                    font-size: 14px;
                }}
                QPushButton:hover {{
                    background-color: {self.colors['primary']}10;
                    border-color: {self.colors['primary_light']};
                }}
                QPushButton:pressed {{
                    background-color: {self.colors['primary']}20;
                }}
            """)
            
    def set_colors(self, colors):
        """Met à jour les couleurs"""
        self.colors = colors
        self.update_style()

class ExcelIntegratedWidget(QWidget):
    """Widget Excel avec intégration forcée et couleurs GULFCAM"""
    
    def __init__(self, parent=None, colors=None):
        super().__init__(parent)
        self.colors = colors or GULFCAM_COLORS
        self.excel = None
        self.workbook = None
        self.current_file = None
        self.is_modified = False
        self.excel_window_hwnd = None
        self.embed_timer = None
        self.retry_count = 0
        self.max_retries = 30
        self.setup_ui()
        
    def setup_ui(self):
        """Configure l'interface Excel avec couleurs GULFCAM"""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        
        # Barre d'outils Excel
        toolbar = QFrame()
        toolbar.setStyleSheet(f"""
            QFrame {{
                background-color: {self.colors['light']};
                border-top-left-radius: 12px;
                border-top-right-radius: 12px;
                border-bottom: 2px solid {self.colors['primary']}20;
                padding: 8px;
            }}
        """)
        toolbar.setFixedHeight(50)
        toolbar_layout = QHBoxLayout(toolbar)
        toolbar_layout.setContentsMargins(12, 8, 12, 8)
        
        # Indicateur de fichier
        self.file_indicator = QLabel("⚡ Aucun fichier chargé")
        self.file_indicator.setStyleSheet(f"""
            QLabel {{
                color: {self.colors['primary']};
                font-weight: 600;
                font-size: 13px;
                padding: 4px 12px;
                background-color: {self.colors['primary']}10;
                border-radius: 16px;
            }}
        """)
        
        # Statut
        self.status_indicator = QLabel("● Prêt")
        self.status_indicator.setStyleSheet(f"""
            QLabel {{
                color: {self.colors['success']};
                font-weight: 600;
                font-size: 12px;
            }}
        """)
        
        # Boutons
        self.save_btn = QPushButton("💾 Enregistrer")
        self.save_btn.setCursor(Qt.PointingHandCursor)
        self.save_btn.clicked.connect(self.save_changes)
        
        self.refresh_btn = QPushButton("🔄 Recharger")
        self.refresh_btn.setCursor(Qt.PointingHandCursor)
        self.refresh_btn.clicked.connect(self.refresh)
        
        self.embed_btn = QPushButton("🔧 Forcer intégration")
        self.embed_btn.setCursor(Qt.PointingHandCursor)
        self.embed_btn.clicked.connect(self.force_embed_excel_now)
        
        # Style des boutons
        for btn in [self.save_btn, self.refresh_btn, self.embed_btn]:
            btn.setStyleSheet(f"""
                QPushButton {{
                    background-color: transparent;
                    color: {self.colors['primary']};
                    border: 1px solid {self.colors['primary']}30;
                    border-radius: 6px;
                    padding: 6px 12px;
                    font-weight: 500;
                    font-size: 12px;
                }}
                QPushButton:hover {{
                    background-color: {self.colors['primary']}10;
                }}
                QPushButton:pressed {{
                    background-color: {self.colors['primary']}20;
                }}
            """)
        
        toolbar_layout.addWidget(self.file_indicator)
        toolbar_layout.addSpacing(20)
        toolbar_layout.addWidget(self.status_indicator)
        toolbar_layout.addStretch()
        toolbar_layout.addWidget(self.save_btn)
        toolbar_layout.addWidget(self.refresh_btn)
        toolbar_layout.addWidget(self.embed_btn)
        
        # Conteneur Excel
        self.excel_container = QFrame()
        self.excel_container.setStyleSheet(f"""
            QFrame {{
                background-color: white;
                border: 2px solid {self.colors['primary']}10;
                border-bottom-left-radius: 12px;
                border-bottom-right-radius: 12px;
                background-color: #f5f5f5;
            }}
        """)
        self.excel_container.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.excel_container.setMinimumHeight(500)
        
        # Layout du conteneur
        container_layout = QVBoxLayout(self.excel_container)
        container_layout.setContentsMargins(10, 10, 10, 10)
        
        # Message placeholder
        self.placeholder = QFrame()
        self.placeholder.setStyleSheet(f"""
            QFrame {{
                background-color: {self.colors['primary']}05;
                border: 2px dashed {self.colors['primary']}30;
                border-radius: 8px;
            }}
        """)
        
        placeholder_layout = QVBoxLayout(self.placeholder)
        placeholder_layout.setAlignment(Qt.AlignCenter)
        placeholder_layout.setSpacing(15)
        
        title_label = QLabel("⚡ GULFCAM Excel Intégré")
        title_label.setStyleSheet(f"color: {self.colors['primary']}; font-size: 20px; font-weight: 600;")
        title_label.setAlignment(Qt.AlignCenter)
        
        desc_label = QLabel("⬇️ Déposez votre fichier Excel ici\nou utilisez le bouton \"Charger\"")
        desc_label.setStyleSheet(f"color: {self.colors['gray']}; font-size: 14px;")
        desc_label.setAlignment(Qt.AlignCenter)
        
        tagline_label = QLabel(GULFCAM_TAGLINE)
        tagline_label.setStyleSheet(f"color: {self.colors['secondary']}; font-size: 12px; font-style: italic;")
        tagline_label.setAlignment(Qt.AlignCenter)
        
        placeholder_layout.addWidget(title_label)
        placeholder_layout.addWidget(desc_label)
        placeholder_layout.addWidget(tagline_label)
        
        container_layout.addWidget(self.placeholder)
        
        layout.addWidget(toolbar)
        layout.addWidget(self.excel_container, 1)
        
        # Timer pour forcer le reparenting
        self.embed_timer = QTimer()
        self.embed_timer.timeout.connect(self.force_embed_excel)
        self.embed_timer.start(1000)
        
    def force_embed_excel_now(self):
        """Force l'intégration d'Excel immédiatement"""
        self.retry_count = 0
        self.force_embed_excel()
        
    def force_embed_excel(self):
        """Force l'intégration d'Excel dans l'application"""
        if self.excel_window_hwnd or not self.excel or not self.workbook:
            return
            
        try:
            self.excel.Visible = True
            time.sleep(0.5)
            
            def find_excel_window(hwnd, ctx):
                try:
                    window_text = win32gui.GetWindowText(hwnd)
                    if self.current_file:
                        file_name = self.current_file.name
                        if file_name in window_text and "Excel" in window_text:
                            ctx['hwnd'] = hwnd
                            return False
                except:
                    pass
                return True
                
            ctx = {'hwnd': None}
            win32gui.EnumWindows(find_excel_window, ctx)
            
            if ctx['hwnd']:
                excel_hwnd = ctx['hwnd']
                
                # Obtenir le handle du conteneur
                container_hwnd = int(self.excel_container.winId())
                
                # Sauvegarder la position originale
                rect = win32gui.GetWindowRect(excel_hwnd)
                
                # Modifier les styles de la fenêtre
                current_style = win32gui.GetWindowLong(excel_hwnd, win32con.GWL_STYLE)
                new_style = current_style & ~(
                    win32con.WS_CAPTION | 
                    win32con.WS_THICKFRAME | 
                    win32con.WS_MINIMIZEBOX | 
                    win32con.WS_MAXIMIZEBOX | 
                    win32con.WS_SYSMENU
                )
                new_style |= win32con.WS_CHILD
                
                win32gui.SetWindowLong(excel_hwnd, win32con.GWL_STYLE, new_style)
                
                # Reparenter la fenêtre
                win32gui.SetParent(excel_hwnd, container_hwnd)
                
                # Redimensionner pour remplir le conteneur
                container_rect = self.excel_container.geometry()
                win32gui.MoveWindow(
                    excel_hwnd, 
                    0, 0, 
                    container_rect.width(), 
                    container_rect.height(), 
                    True
                )
                
                self.excel_window_hwnd = excel_hwnd
                self.placeholder.hide()
                
                # Rafraîchir
                win32gui.InvalidateRect(container_hwnd, None, True)
                win32gui.UpdateWindow(container_hwnd)
                
                print(f"✅ Excel intégré avec succès: {excel_hwnd}")
                self.retry_count = 0
                self.embed_timer.stop()
                
            else:
                self.retry_count += 1
                if self.retry_count > self.max_retries:
                    print("❌ Impossible de trouver la fenêtre Excel")
                    self.status_indicator.setText("● Échec intégration")
                    self.status_indicator.setStyleSheet(f"color: {self.colors['danger']}; font-weight: 600;")
                    
        except Exception as e:
            print(f"⚠️ Erreur embedding Excel: {e}")
            self.retry_count += 1
            if self.retry_count > self.max_retries:
                self.embed_timer.stop()
                
    def resize_excel_window(self):
        """Redimensionne Excel pour remplir le conteneur"""
        if self.excel_window_hwnd:
            try:
                container_rect = self.excel_container.geometry()
                win32gui.MoveWindow(
                    self.excel_window_hwnd,
                    0, 0,
                    container_rect.width(),
                    container_rect.height(),
                    True
                )
                win32gui.InvalidateRect(self.excel_window_hwnd, None, True)
            except Exception as e:
                print(f"Erreur resize: {e}")
            
    def resizeEvent(self, event):
        """Redimensionne Excel quand le widget change de taille"""
        super().resizeEvent(event)
        self.resize_excel_window()
        
    def showEvent(self, event):
        """Quand le widget devient visible"""
        super().showEvent(event)
        QTimer.singleShot(100, self.resize_excel_window)
        
    def load_file(self, file_path: Path) -> bool:
        """Charge un fichier Excel"""
        try:
            if not file_path or not file_path.exists():
                return False
                
            self.current_file = file_path
            self.file_indicator.setText(f"⚡ {file_path.name}")
            
            # Initialiser Excel
            if self.excel is None:
                self.excel = win32com.client.Dispatch("Excel.Application")
                self.excel.DisplayAlerts = False
                
            # Fermer l'ancien classeur
            if self.workbook:
                try:
                    self.workbook.Close(SaveChanges=False)
                except:
                    pass
                    
            # Ouvrir le nouveau fichier
            self.excel.Visible = True
            self.workbook = self.excel.Workbooks.Open(str(file_path.absolute()))
            
            # Réinitialiser
            if self.excel_window_hwnd:
                try:
                    win32gui.DestroyWindow(self.excel_window_hwnd)
                except:
                    pass
                    
            self.excel_window_hwnd = None
            self.retry_count = 0
            self.placeholder.show()
            
            # Forcer l'embedding
            self.embed_timer.start(1000)
            self.force_embed_excel()
            
            self.status_indicator.setText("● Chargé")
            self.status_indicator.setStyleSheet(f"color: {self.colors['success']}; font-weight: 600;")
            
            return True
            
        except Exception as e:
            QMessageBox.critical(self, "Erreur GULFCAM", f"Impossible de charger le fichier:\n{str(e)}")
            return False
            
    def save_changes(self):
        """Enregistre les modifications"""
        if self.workbook:
            try:
                self.workbook.Save()
                self.status_indicator.setText("● Enregistré")
                QMessageBox.information(self, "GULFCAM", "Modifications enregistrées avec succès!")
            except Exception as e:
                QMessageBox.critical(self, "Erreur GULFCAM", f"Impossible d'enregistrer:\n{str(e)}")
                
    def refresh(self):
        """Recharge le fichier"""
        if self.current_file and self.current_file.exists():
            reply = QMessageBox.question(
                self, "Confirmation GULFCAM",
                "Recharger le fichier ? Les modifications non enregistrées seront perdues.",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply == QMessageBox.Yes:
                self.load_file(self.current_file)
                
    def closeEvent(self, event):
        """Fermeture propre"""
        self.embed_timer.stop()
        if self.workbook:
            try:
                self.workbook.Close(SaveChanges=False)
            except:
                pass
        if self.excel:
            try:
                self.excel.Quit()
            except:
                pass
        event.accept()

class CompanyDashboard(QWidget):
    """Dashboard entreprise avec branding GULFCAM"""
    
    def __init__(self, colors):
        super().__init__()
        self.colors = colors
        self.setup_ui()
        
    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(24)
        
        # En-tête avec logo
        header = QFrame()
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(0, 0, 0, 0)
        
        # Logo GULFCAM avec tagline
        logo_container = QFrame()
        logo_layout = QHBoxLayout(logo_container)
        logo_layout.setSpacing(16)
        
        logo_text = QLabel("GULFCAM")
        logo_text.setStyleSheet(f"""
            font-size: 28px;
            font-weight: 800;
            color: {self.colors['primary']};
            letter-spacing: 2px;
        """)
        
        tagline = QLabel(GULFCAM_TAGLINE)
        tagline.setStyleSheet(f"""
            color: {self.colors['secondary']};
            font-size: 14px;
            font-style: italic;
            padding-top: 8px;
        """)
        
        logo_layout.addWidget(logo_text)
        logo_layout.addWidget(tagline)
        
        date_label = QLabel(time.strftime("%d %B %Y"))
        date_label.setStyleSheet(f"""
            color: {self.colors['gray']};
            font-size: 14px;
            padding: 8px 16px;
            background-color: {self.colors['light-gray']};
            border-radius: 20px;
        """)
        
        header_layout.addWidget(logo_container)
        header_layout.addStretch()
        header_layout.addWidget(date_label)
        
        layout.addWidget(header)
        
        # Message de bienvenue
        welcome_card = ModernCard(colors=self.colors)
        welcome_layout = QHBoxLayout(welcome_card)
        
        welcome_icon = QLabel("⚡")
        welcome_icon.setStyleSheet(f"font-size: 32px; color: {self.colors['secondary']};")
        
        welcome_text = QLabel(f"Bienvenue sur la plateforme DSF de {GULFCAM_LOGO_TEXT}")
        welcome_text.setStyleSheet(f"""
            font-size: 20px;
            font-weight: 600;
            color: {self.colors['primary']};
        """)
        
        welcome_layout.addWidget(welcome_icon)
        welcome_layout.addWidget(welcome_text)
        welcome_layout.addStretch()
        
        layout.addWidget(welcome_card)
        
        # Cartes KPI
        kpi_grid = QGridLayout()
        kpi_grid.setSpacing(20)
        
        kpis = [
            ("📊 Documents traités", "24", "+12%", self.colors['success']),
            ("⏱️ Temps moyen", "3.2 min", "-8%", self.colors['primary']),
            ("✅ Taux de réussite", "98.5%", "+2.3%", self.colors['success']),
            ("📈 Projets actifs", "8", "0", self.colors['secondary']),
        ]
        
        for i, (title, value, trend, color) in enumerate(kpis):
            card = ModernCard(colors=self.colors)
            card_layout = QVBoxLayout(card)
            
            title_label = QLabel(title)
            title_label.setStyleSheet(f"color: {self.colors['gray']}; font-size: 14px;")
            
            value_label = QLabel(value)
            value_label.setStyleSheet(f"""
                font-size: 36px;
                font-weight: 700;
                color: {self.colors['dark']};
            """)
            
            trend_label = QLabel(trend)
            trend_label.setStyleSheet(f"""
                color: {color};
                font-weight: 600;
                font-size: 13px;
                padding: 4px 8px;
                background-color: {color}20;
                border-radius: 12px;
            """)
            
            card_layout.addWidget(title_label)
            card_layout.addWidget(value_label)
            card_layout.addWidget(trend_label, 0, Qt.AlignRight)
            
            kpi_grid.addWidget(card, i // 2, i % 2)
            
        layout.addLayout(kpi_grid)
        
        # Actions rapides
        actions_card = ModernCard(colors=self.colors)
        actions_layout = QVBoxLayout(actions_card)
        
        actions_title = QLabel("Actions rapides GULFCAM")
        actions_title.setStyleSheet(f"""
            font-size: 18px;
            font-weight: 600;
            color: {self.colors['primary']};
            margin-bottom: 16px;
        """)
        
        actions_grid = QGridLayout()
        actions_grid.setSpacing(12)
        
        actions = [
            ("Nouveau DSF", "📄", self.colors['primary']),
            ("Charger balance", "📤", self.colors['success']),
            ("Rapports", "📊", self.colors['secondary']),
            ("Paramètres", "⚙️", self.colors['gray']),
        ]
        
        for i, (text, icon, color) in enumerate(actions):
            btn = QPushButton(f"  {icon}  {text}")
            btn.setCursor(Qt.PointingHandCursor)
            btn.setStyleSheet(f"""
                QPushButton {{
                    background-color: {color}10;
                    color: {color};
                    border: 1px solid {color}30;
                    border-radius: 8px;
                    padding: 16px;
                    font-weight: 600;
                    text-align: left;
                    font-size: 14px;
                }}
                QPushButton:hover {{
                    background-color: {color}20;
                }}
            """)
                
            actions_grid.addWidget(btn, i // 2, i % 2)
            
        actions_layout.addWidget(actions_title)
        actions_layout.addLayout(actions_grid)
        
        layout.addWidget(actions_card)

class MainWorkView(QWidget):
    """Vue de travail principale avec branding GULFCAM"""
    
    def __init__(self, controller, colors):
        super().__init__()
        self.controller = controller
        self.colors = colors
        self.balance_file = None
        self.setup_ui()
        
    def setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        # Barre de navigation
        nav_bar = QFrame()
        nav_bar.setStyleSheet(f"""
            QFrame {{
                background-color: white;
                border-bottom: 2px solid {self.colors['primary']}20;
                padding: 8px 24px;
            }}
        """)
        nav_layout = QHBoxLayout(nav_bar)
        nav_layout.setContentsMargins(0, 0, 0, 0)
        
        # Logo GULFCAM
        logo_container = QFrame()
        logo_layout = QHBoxLayout(logo_container)
        logo_layout.setSpacing(8)
        
        logo_text = QLabel("G")
        logo_text.setStyleSheet(f"""
            font-size: 24px;
            font-weight: 800;
            color: white;
            background-color: {self.colors['primary']};
            padding: 8px 16px;
            border-radius: 8px;
        """)
        
        logo_full = QLabel("GULFCAM DSF")
        logo_full.setStyleSheet(f"""
            font-size: 18px;
            font-weight: 700;
            color: {self.colors['primary']};
            letter-spacing: 1px;
        """)
        
        logo_layout.addWidget(logo_text)
        logo_layout.addWidget(logo_full)
        
        # Navigation
        nav_buttons = [
            ("🏠 Accueil", 0),
            ("📊 Balance", 1),
            ("📝 DSF", 2),
            ("📈 Rapports", 3),
        ]
        
        nav_layout.addWidget(logo_container)
        nav_layout.addSpacing(40)
        
        for text, idx in nav_buttons:
            btn = QPushButton(text)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setStyleSheet(f"""
                QPushButton {{
                    background-color: transparent;
                    color: {self.colors['gray']};
                    border: none;
                    padding: 8px 16px;
                    font-weight: 500;
                    border-radius: 20px;
                }}
                QPushButton:hover {{
                    background-color: {self.colors['primary']}10;
                    color: {self.colors['primary']};
                }}
                QPushButton:pressed {{
                    background-color: {self.colors['secondary']}20;
                    color: {self.colors['secondary']};
                }}
            """)
            btn.clicked.connect(lambda checked, x=idx: self.controller.switch_view(x))
            nav_layout.addWidget(btn)
            
        nav_layout.addStretch()
        
        # Thème toggle
        theme_btn = QPushButton("🌓")
        theme_btn.setCursor(Qt.PointingHandCursor)
        theme_btn.setFixedSize(40, 40)
        theme_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {self.colors['primary']}10;
                color: {self.colors['primary']};
                border: none;
                border-radius: 20px;
                font-size: 18px;
            }}
            QPushButton:hover {{
                background-color: {self.colors['primary']};
                color: white;
            }}
        """)
        theme_btn.clicked.connect(self.controller.toggle_dark_mode)
        nav_layout.addWidget(theme_btn)
        
        main_layout.addWidget(nav_bar)
        
        # Zone de contenu avec scroll
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("""
            QScrollArea {
                border: none;
                background-color: transparent;
            }
        """)
        
        content = QWidget()
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(24, 24, 24, 24)
        content_layout.setSpacing(24)
        
        # Dashboard
        self.dashboard = CompanyDashboard(self.colors)
        content_layout.addWidget(self.dashboard)
        
        # Section Balance
        balance_card = ModernCard(colors=self.colors)
        balance_layout = QVBoxLayout(balance_card)
        
        balance_title = QLabel("1. Chargement de la balance")
        balance_title.setStyleSheet(f"""
            font-size: 20px;
            font-weight: 600;
            color: {self.colors['primary']};
            margin-bottom: 16px;
        """)
        
        # Zone de sélection de fichier
        file_row = QHBoxLayout()
        
        self.balance_path = QLineEdit()
        self.balance_path.setPlaceholderText("Sélectionnez le fichier balance (.xlsx)")
        self.balance_path.setStyleSheet(f"""
            QLineEdit {{
                border: 2px solid {self.colors['light-gray']};
                border-radius: 8px;
                padding: 12px;
                font-size: 14px;
                background-color: {self.colors['card-bg']};
                color: {self.colors['dark']};
            }}
            QLineEdit:focus {{
                border-color: {self.colors['secondary']};
            }}
        """)
        
        browse_btn = AnimatedButton("Parcourir", primary=False, colors=self.colors)
        load_btn = AnimatedButton("Charger", primary=True, colors=self.colors)
        load_btn.clicked.connect(self.load_balance)
        
        browse_btn.clicked.connect(self.pick_balance)
        
        file_row.addWidget(self.balance_path, 1)
        file_row.addWidget(browse_btn)
        file_row.addWidget(load_btn)
        
        balance_layout.addWidget(balance_title)
        balance_layout.addLayout(file_row)
        
        content_layout.addWidget(balance_card)
        
        # Section Excel intégré
        excel_card = ModernCard(colors=self.colors)
        excel_layout = QVBoxLayout(excel_card)
        
        excel_title = QLabel("2. Visualisation et édition Excel")
        excel_title.setStyleSheet(f"""
            font-size: 20px;
            font-weight: 600;
            color: {self.colors['primary']};
            margin-bottom: 16px;
        """)
        
        excel_layout.addWidget(excel_title)
        
        # Widget Excel intégré
        self.excel_widget = ExcelIntegratedWidget(colors=self.colors)
        excel_layout.addWidget(self.excel_widget, 1)
        
        content_layout.addWidget(excel_card)
        
        # Section DSF
        dsf_card = ModernCard(colors=self.colors)
        dsf_layout = QVBoxLayout(dsf_card)
        
        dsf_title = QLabel("3. Génération DSF")
        dsf_title.setStyleSheet(f"""
            font-size: 20px;
            font-weight: 600;
            color: {self.colors['primary']};
            margin-bottom: 16px;
        """)
        
        dsf_actions = QHBoxLayout()
        
        generate_btn = AnimatedButton("Générer le DSF", primary=True, colors=self.colors)
        generate_btn.clicked.connect(self.generate_dsf)
        
        export_btn = AnimatedButton("Exporter", primary=False, colors=self.colors)
        export_btn.clicked.connect(self.export_dsf)
        
        dsf_actions.addWidget(generate_btn)
        dsf_actions.addWidget(export_btn)
        dsf_actions.addStretch()
        
        dsf_layout.addWidget(dsf_title)
        dsf_layout.addLayout(dsf_actions)
        
        content_layout.addWidget(dsf_card)
        
        # Footer
        footer = QLabel(f"{GULFCAM_LOGO_TEXT} • {GULFCAM_TAGLINE} • DSF Mapper v1.0")
        footer.setAlignment(Qt.AlignCenter)
        footer.setStyleSheet(f"""
            color: {self.colors['gray']};
            font-size: 12px;
            padding: 16px;
            border-top: 1px solid {self.colors['light-gray']};
        """)
        content_layout.addWidget(footer)
        
        scroll.setWidget(content)
        main_layout.addWidget(scroll, 1)
        
    def pick_balance(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Sélectionner le bilan - GULFCAM",
            str(Path.home() / "Desktop"),
            "Fichiers Excel (*.xlsx)"
        )
        if path:
            self.balance_path.setText(path)
            
    def load_balance(self):
        path = self.balance_path.text().strip()
        if not path or not Path(path).exists():
            QMessageBox.warning(self, "Erreur GULFCAM", "Veuillez sélectionner un fichier valide")
            return
            
        if self.excel_widget.load_file(Path(path)):
            self.balance_file = path
            self.controller.statusBar().showMessage(f"⚡ GULFCAM - Balance chargée: {Path(path).name}")
            
    def generate_dsf(self):
        if not self.balance_file:
            QMessageBox.warning(self, "Erreur GULFCAM", "Veuillez d'abord charger une balance")
            return
            
        QMessageBox.information(self, "GULFCAM", "Fonctionnalité de génération DSF à implémenter")
        
    def export_dsf(self):
        if not self.excel_widget.workbook:
            QMessageBox.warning(self, "Erreur GULFCAM", "Aucun fichier à exporter")
            return
            
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Exporter le DSF - GULFCAM",
            str(Path.home() / "Desktop"),
            "Fichiers Excel (*.xlsx)"
        )
        
        if path:
            if not path.endswith('.xlsx'):
                path += '.xlsx'
            try:
                self.excel_widget.save_changes()
                shutil.copy(str(self.excel_widget.current_file), path)
                QMessageBox.information(self, "GULFCAM", f"Fichier exporté avec succès vers:\n{path}")
            except Exception as e:
                QMessageBox.critical(self, "Erreur GULFCAM", f"Export impossible:\n{str(e)}")
                
    def set_colors(self, colors):
        """Met à jour les couleurs"""
        self.colors = colors
        self.dashboard.set_colors(colors)
        self.excel_widget.set_colors(colors)

class LightDSFMapperWindow(QMainWindow):
    """Fenêtre principale avec branding GULFCAM"""
    
    def __init__(self):
        super().__init__()
        
        # Charger les paramètres
        self.settings = AppSettings.load()
        self.colors = get_gulfcam_colors(self.settings.dark_mode)
        
        self.setWindowTitle("GULFCAM DSF Mapper")
        self.setWindowIcon(self.create_gulfcam_icon())
        self.resize(self.settings.window_width, self.settings.window_height)
        
        # Widget central
        central = QWidget()
        self.setCentralWidget(central)
        
        # Stack pour les différentes vues
        self.stack = QStackedWidget()
        layout = QVBoxLayout(central)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.stack)
        
        # Vue de travail principale
        self.main_view = MainWorkView(self, self.colors)
        self.stack.addWidget(self.main_view)
        
        # Barre d'état
        self.setup_status_bar()
        
    def create_gulfcam_icon(self):
        """Crée une icône pour l'application"""
        icon = QIcon()
        pixmap = QPixmap(64, 64)
        pixmap.fill(QColor(self.colors['primary']))
        
        painter = QPainter(pixmap)
        painter.setPen(Qt.white)
        painter.setFont(QFont("Arial", 24, QFont.Bold))
        painter.drawText(pixmap.rect(), Qt.AlignCenter, "G")
        painter.end()
        
        icon.addPixmap(pixmap)
        return icon
        
    def setup_status_bar(self):
        status_bar = QStatusBar()
        self.setStatusBar(status_bar)
        status_bar.setStyleSheet(f"""
            QStatusBar {{
                background-color: {self.colors['primary']};
                color: white;
                padding: 8px;
                font-size: 12px;
            }}
        """)
        status_bar.showMessage(f"⚡ {GULFCAM_LOGO_TEXT} • {GULFCAM_TAGLINE} • Application prête")
        
    def switch_view(self, index):
        """Change la vue active"""
        self.stack.setCurrentIndex(index)
        
    def toggle_dark_mode(self):
        """Bascule le mode sombre/clair"""
        self.settings.dark_mode = not self.settings.dark_mode
        self.colors = get_gulfcam_colors(self.settings.dark_mode)
        self.settings.save()
        
        # Rafraîchir l'interface
        self.main_view = MainWorkView(self, self.colors)
        self.stack.removeWidget(self.stack.currentWidget())
        self.stack.addWidget(self.main_view)
        
        # Mettre à jour la barre d'état
        self.statusBar().setStyleSheet(f"""
            QStatusBar {{
                background-color: {self.colors['primary']};
                color: white;
                padding: 8px;
                font-size: 12px;
            }}
        """)
        
        self.statusBar().showMessage(f"🌓 Mode {'sombre' if self.settings.dark_mode else 'clair'} GULFCAM activé")
        
    def closeEvent(self, event):
        """Sauvegarde les paramètres à la fermeture"""
        self.settings.window_width = self.width()
        self.settings.window_height = self.height()
        self.settings.save()
        
        # Fermer Excel proprement
        if hasattr(self.main_view, 'excel_widget'):
            self.main_view.excel_widget.closeEvent(event)
            
        event.accept()

def main():
    """Point d'entrée principal"""
    # Initialiser COM
    pythoncom.CoInitialize()
    
    app = QApplication([])
    
    # Police par défaut
    font = QFont("Segoe UI", 10)
    app.setFont(font)
    
    # Variable pour stocker la fenêtre principale
    main_window = None
    
    def show_main_window():
        nonlocal main_window
        main_window = LightDSFMapperWindow()
        main_window.show()
    
    # Afficher le splash screen
    splash = SplashScreen(show_main_window)
    splash.show()
    
    result = app.exec()
    
    # Nettoyer COM
    pythoncom.CoUninitialize()
    
    return result

if __name__ == "__main__":
    main()