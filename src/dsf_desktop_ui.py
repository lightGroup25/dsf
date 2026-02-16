#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""LIGHT DSF MAPPER – Design GULFCAM avec Landing Page Animée et Excel Intégré"""

from __future__ import annotations

import json
import math
import os
import shutil
import subprocess
import time
import tempfile
from pathlib import Path
from typing import Dict, List, Optional
from dataclasses import dataclass, asdict

IS_WINDOWS = os.name == "nt"

if IS_WINDOWS:
    try:
        import pythoncom
        import win32com.client
        import win32gui
        import win32con
        WINDOWS_COM_AVAILABLE = True
    except ImportError:
        pythoncom = None
        win32com = None
        win32gui = None
        win32con = None
        WINDOWS_COM_AVAILABLE = False
else:
    pythoncom = None
    win32com = None
    win32gui = None
    win32con = None
    WINDOWS_COM_AVAILABLE = False

from PySide6.QtCore import Qt, QTimer, QSize, QPropertyAnimation, QEasingCurve
from PySide6.QtGui import QFont, QAction, QColor, QIcon, QKeySequence, QShortcut, QPixmap, QPainter
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QFileDialog,
    QHeaderView,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSizePolicy,
    QStackedWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
    QFrame,
    QProgressBar,
    QGraphicsDropShadowEffect,
    QStatusBar,
    QStyle,
    QGridLayout,
    QScrollArea,
)

# FontAwesome (qtawesome)
try:
    import qtawesome as qta
    FONTS_AVAILABLE = True
except ImportError:
    qta = None
    FONTS_AVAILABLE = False

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
    'primary': '#118E55',        # Vert renforcé pour meilleure visibilité
    'primary_light': '#1CAC69',  # Vert lumineux pour hover/active
    'secondary': '#D8173A',      # Rouge logo (anneau)
    'secondary_light': '#EF3E5D',# Rouge adouci
    'accent': '#F0C419',         # Jaune logo (swoosh)
    'success': '#28B273',        # Vert de confirmation
    'danger': '#C81E35',         # Rouge d'erreur
    'warning': '#D8A90E',        # Jaune d'alerte
    'light': '#F3F9F4',          # Fond clair légèrement verdoyant
    'dark': '#102D21',           # Texte sombre lisible
    'gray': '#5E7166',           # Gris chaud neutre
    'light-gray': '#D6E3DA',     # Bordures claires
    'card-bg': '#FFFFFF',        # Fond de carte
    'gradient_start': '#11784A', # Dégradé splash/header
    'gradient_end': '#1CAC69',
}

# Pour le mode sombre
GULFCAM_DARK_COLORS = {
    'primary': '#4DD38E',        # Vert renforcé pour fond sombre
    'primary_light': '#73E8AB',
    'secondary': '#FF5A74',      # Rouge réhaussé pour contraste
    'secondary_light': '#FF8095',
    'accent': '#F7D14A',         # Jaune réchauffé
    'success': '#44E0A5',
    'danger': '#FB7185',
    'warning': '#F4C74F',
    'light': '#101A14',          # Fond global sombre
    'dark': '#F2F7F3',           # Texte principal clair
    'gray': '#A6B8AE',           # Texte secondaire
    'light-gray': '#2A3D33',     # Bordures sombres
    'card-bg': '#18251E',        # Cartes en mode sombre
    'gradient_start': '#0E2A1B',
    'gradient_end': '#1E6B47',
}

# Logo GULFCAM
GULFCAM_LOGO_TEXT = "GULFCAM"
GULFCAM_TAGLINE = "S'investir pour vous"
PRIMARY_LOGO_FILENAME = "OIP.png"
PRIMARY_LOGO_FALLBACK = "OIP.webp"
LIGHTGROUP_LOGO_FILENAME = "lightgroup.png"

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


def resolve_image_path(filename: str) -> Optional[Path]:
    """Résout le chemin d'une image dans le projet."""
    here = Path(__file__).resolve().parent
    candidates = (
        Path.cwd() / filename,
        here / filename,
        here.parent / filename,
    )
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return None


def load_image_pixmap(filename: str, width: int, height: int) -> QPixmap:
    """Charge une image du projet, redimensionnée proprement."""
    image_path = resolve_image_path(filename)
    if image_path is None:
        return QPixmap()
    pixmap = QPixmap(str(image_path))
    if pixmap.isNull():
        return QPixmap()
    return pixmap.scaled(width, height, Qt.KeepAspectRatio, Qt.SmoothTransformation)


def load_logo_pixmap(width: int, height: int) -> QPixmap:
    """Charge le logo principal (OIP.png), puis fallback."""
    pixmap = load_image_pixmap(PRIMARY_LOGO_FILENAME, width, height)
    if not pixmap.isNull():
        return pixmap
    return load_image_pixmap(PRIMARY_LOGO_FALLBACK, width, height)


def load_lightgroup_pixmap(width: int, height: int) -> QPixmap:
    """Charge le logo Light Group."""
    return load_image_pixmap(LIGHTGROUP_LOGO_FILENAME, width, height)


def ui_icon(
    widget: QWidget,
    fa_name: str,
    color: Optional[str] = None,
    fallback: Optional[QStyle.StandardPixmap] = None,
) -> QIcon:
    """Retourne une icône FontAwesome (qtawesome), ou une fallback Qt."""
    if FONTS_AVAILABLE and qta is not None:
        kwargs = {}
        if color:
            kwargs["color"] = color
        try:
            return qta.icon(fa_name, **kwargs)
        except Exception:
            pass
    if fallback is not None:
        return widget.style().standardIcon(fallback)
    return QIcon()

class SplashScreen(QWidget):
    """Splash screen avec progression réelle des étapes de démarrage."""

    def __init__(self, callback, startup_tasks=None, min_duration_ms: int = 5000):
        super().__init__()
        self.callback = callback
        self.startup_tasks = startup_tasks or []
        self.min_duration_ms = max(2000, min_duration_ms)
        self.current_task_index = 0
        self.started_at_ms = 0.0
        self.task_failed = False

        self.splash_logo_label: Optional[QLabel] = None
        self.splash_logo_base = QPixmap()
        self.logo_glow_effect: Optional[QGraphicsDropShadowEffect] = None
        self.logo_glow_timer = QTimer(self)
        self.logo_glow_timer.setInterval(40)
        self.logo_glow_timer.timeout.connect(self._animate_logo_glow)
        self.logo_glow_elapsed_ms = 0

        self.progress_bar_width = 360
        self.progress_value = 0
        self.progress_fill: Optional[QFrame] = None
        self.loading_label: Optional[QLabel] = None

        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setAttribute(Qt.WA_TranslucentBackground)

        screen = QApplication.primaryScreen().geometry()
        self.setGeometry(screen)

        self.opacity_animation = QPropertyAnimation(self, b"windowOpacity")
        self.opacity_animation.setDuration(700)
        self.opacity_animation.setStartValue(0.0)
        self.opacity_animation.setEndValue(1.0)
        self.opacity_animation.setEasingCurve(QEasingCurve.OutCubic)

        self.setup_ui()
        self.start_animations()

    def setup_ui(self):
        """Configure l'interface du splash screen."""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setAlignment(Qt.AlignCenter)

        container = QFrame()
        container.setStyleSheet(f"""
            QFrame {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                    stop:0 #0e2b1d,
                    stop:0.55 #1a5b3d,
                    stop:1 #2d7455);
                border-radius: 0px;
            }}
        """)
        container.setFixedSize(self.width(), self.height())

        container_layout = QVBoxLayout(container)
        container_layout.setAlignment(Qt.AlignCenter)
        container_layout.setSpacing(18)

        logo_container = QFrame()
        logo_layout = QVBoxLayout(logo_container)
        logo_layout.setAlignment(Qt.AlignCenter)
        logo_layout.setSpacing(10)

        logo_card = QFrame()
        logo_card.setStyleSheet("""
            QFrame {
                background-color: rgba(255,255,255,0.12);
                border: 1px solid rgba(255,255,255,0.28);
                border-radius: 20px;
                padding: 10px;
            }
        """)
        logo_card_layout = QVBoxLayout(logo_card)
        logo_card_layout.setContentsMargins(14, 14, 14, 14)
        logo_card_layout.setAlignment(Qt.AlignCenter)

        self.splash_logo_base = load_logo_pixmap(185, 195)
        self.splash_logo_label = QLabel()
        self.splash_logo_label.setFixedSize(185, 195)
        self.splash_logo_label.setAlignment(Qt.AlignCenter)

        if not self.splash_logo_base.isNull():
            self.splash_logo_label.setPixmap(self.splash_logo_base)
            self.splash_logo_label.setStyleSheet("background: transparent;")
            self.logo_glow_effect = QGraphicsDropShadowEffect(self.splash_logo_label)
            self.logo_glow_effect.setOffset(0, 0)
            self.logo_glow_effect.setBlurRadius(26)
            glow_color = QColor("#f0c419")
            glow_color.setAlpha(130)
            self.logo_glow_effect.setColor(glow_color)
            self.splash_logo_label.setGraphicsEffect(self.logo_glow_effect)
        else:
            self.splash_logo_label.setText(GULFCAM_LOGO_TEXT)
            self.splash_logo_label.setStyleSheet("""
                QLabel {
                    color: white;
                    font-size: 56px;
                    font-weight: 800;
                    letter-spacing: 6px;
                    background: transparent;
                }
            """)

        tagline_label = QLabel(GULFCAM_TAGLINE)
        tagline_label.setStyleSheet("""
            QLabel {
                color: rgba(255,255,255,0.95);
                font-size: 24px;
                font-weight: 700;
                letter-spacing: 1px;
                background: transparent;
            }
        """)
        tagline_label.setAlignment(Qt.AlignCenter)

        logo_card_layout.addWidget(self.splash_logo_label, 0, Qt.AlignCenter)
        logo_layout.addWidget(logo_card, 0, Qt.AlignCenter)
        logo_layout.addWidget(tagline_label, 0, Qt.AlignCenter)

        footer_widget = QWidget()
        footer_layout = QHBoxLayout(footer_widget)
        footer_layout.setContentsMargins(0, 0, 0, 0)
        footer_layout.setSpacing(12)
        footer_layout.setAlignment(Qt.AlignCenter)

        dsf_by_label = QLabel("DSF by")
        dsf_by_label.setStyleSheet("""
            QLabel {
                color: rgba(255,255,255,0.98);
                font-size: 34px;
                font-weight: 700;
                letter-spacing: 2px;
                background: transparent;
            }
        """)

        lightgroup_badge = QFrame()
        lightgroup_badge.setStyleSheet("""
            QFrame {
                background-color: rgba(255,255,255,0.18);
                border: 1px solid rgba(255,255,255,0.30);
                border-radius: 12px;
            }
        """)
        lightgroup_badge_layout = QHBoxLayout(lightgroup_badge)
        lightgroup_badge_layout.setContentsMargins(10, 6, 10, 6)
        lightgroup_badge_layout.setAlignment(Qt.AlignCenter)

        lightgroup_footer_logo = QLabel()
        lightgroup_footer_logo.setAlignment(Qt.AlignCenter)
        lightgroup_footer_pixmap = load_lightgroup_pixmap(150, 60)
        if not lightgroup_footer_pixmap.isNull():
            lightgroup_footer_logo.setPixmap(lightgroup_footer_pixmap)
        else:
            lightgroup_footer_logo.setText("LIGHTGROUP")
            lightgroup_footer_logo.setStyleSheet("""
                QLabel {
                    color: rgba(255,255,255,0.95);
                    font-size: 20px;
                    font-weight: 700;
                    letter-spacing: 2px;
                    background: transparent;
                }
            """)
        lightgroup_badge_layout.addWidget(lightgroup_footer_logo)

        footer_layout.addWidget(dsf_by_label)
        footer_layout.addWidget(lightgroup_badge)

        self.progress_bar = QFrame()
        self.progress_bar.setFixedSize(self.progress_bar_width, 6)
        self.progress_bar.setStyleSheet("""
            QFrame {
                background-color: rgba(255,255,255,0.22);
                border-radius: 3px;
            }
        """)

        self.progress_fill = QFrame(self.progress_bar)
        self.progress_fill.setFixedSize(0, 6)
        self.progress_fill.setStyleSheet(f"""
            QFrame {{
                background-color: {GULFCAM_COLORS['accent']};
                border-radius: 3px;
            }}
        """)

        self.loading_label = QLabel("Initialisation...")
        self.loading_label.setStyleSheet("""
            QLabel {
                color: rgba(255,255,255,0.82);
                font-size: 15px;
                font-weight: 500;
                background: transparent;
            }
        """)
        self.loading_label.setAlignment(Qt.AlignCenter)

        container_layout.addWidget(logo_container, 0, Qt.AlignCenter)
        container_layout.addSpacing(24)
        container_layout.addWidget(footer_widget, 0, Qt.AlignCenter)
        container_layout.addSpacing(44)
        container_layout.addWidget(self.progress_bar, 0, Qt.AlignCenter)
        container_layout.addWidget(self.loading_label, 0, Qt.AlignCenter)

        layout.addWidget(container)

    def start_animations(self):
        """Démarre les animations et le chargement réel."""
        self.opacity_animation.start()
        self.logo_glow_elapsed_ms = 0
        self.logo_glow_timer.start()
        self.started_at_ms = time.time() * 1000.0
        self._set_progress(3)
        QTimer.singleShot(80, self._run_next_task)

    def _set_progress(self, value: int):
        """Met à jour la barre de progression."""
        self.progress_value = max(0, min(100, value))
        fill_width = int((self.progress_bar_width * self.progress_value) / 100)
        if self.progress_fill is not None:
            self.progress_fill.setGeometry(0, 0, fill_width, 6)

    def _run_next_task(self):
        """Exécute les tâches de démarrage en séquence."""
        if self.task_failed:
            return

        if self.current_task_index >= len(self.startup_tasks):
            self._finalize_startup()
            return

        step_name, step_fn = self.startup_tasks[self.current_task_index]
        if self.loading_label is not None:
            self.loading_label.setText(step_name)
        QApplication.processEvents()

        try:
            step_fn()
        except Exception as exc:
            self.task_failed = True
            if self.loading_label is not None:
                self.loading_label.setText(f"Erreur au démarrage: {exc}")
            QTimer.singleShot(1200, self.close_splash)
            return

        self.current_task_index += 1
        task_ratio = self.current_task_index / max(1, len(self.startup_tasks))
        self._set_progress(int(5 + (task_ratio * 90)))
        QTimer.singleShot(60, self._run_next_task)

    def _finalize_startup(self):
        """Attend la durée minimale (5s), puis termine à 100%."""
        elapsed = (time.time() * 1000.0) - self.started_at_ms
        remaining = max(0, int(self.min_duration_ms - elapsed))

        if self.loading_label is not None:
            self.loading_label.setText("Chargement de l'application...")

        if remaining > 0:
            self._set_progress(97)
            QTimer.singleShot(remaining, self._finish_and_close)
        else:
            self._finish_and_close()

    def _finish_and_close(self):
        self._set_progress(100)
        if self.loading_label is not None:
            self.loading_label.setText("Application prête")
        QTimer.singleShot(150, self.close_splash)

    def _animate_logo_glow(self):
        """Fait varier un halo lumineux autour du logo."""
        self.logo_glow_elapsed_ms += self.logo_glow_timer.interval()
        if self.logo_glow_effect is None:
            return

        period_ms = 1050.0
        wave = (math.sin((2.0 * math.pi * self.logo_glow_elapsed_ms) / period_ms) + 1.0) / 2.0
        blur = 18.0 + (22.0 * wave)
        alpha = int(70 + (95 * wave))

        glow_color = QColor(GULFCAM_COLORS["accent"])
        glow_color.setAlpha(alpha)
        self.logo_glow_effect.setBlurRadius(blur)
        self.logo_glow_effect.setColor(glow_color)
        self.logo_glow_effect.setOffset(0, 0)

    def close_splash(self):
        """Ferme le splash screen et affiche l'application."""
        if self.logo_glow_timer.isActive():
            self.logo_glow_timer.stop()
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
        shadow_color = QColor(self.colors['primary'])
        shadow_color.setAlpha(32)
        shadow.setColor(shadow_color)
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
        self.preview_workbook = None
        self.preview_sheet_name = None
        self.current_file = None
        self.is_modified = False
        self.excel_window_hwnd = None
        self.embed_timer = None
        self.retry_count = 0
        self.max_retries = 30
        self.preview_max_rows = 350
        self.preview_max_cols = 80
        self._updating_table = False
        self.status_color_key = "success"
        self.use_windows_com = WINDOWS_COM_AVAILABLE
        self.spreadsheet_command = self._detect_spreadsheet_command()
        self.external_editor_process = None
        self.setup_ui()

    def _detect_spreadsheet_command(self) -> Optional[List[str]]:
        """Détecte un tableur disponible sur la machine."""
        for binary in ("libreoffice", "soffice"):
            if shutil.which(binary):
                return [binary, "--calc"]
        return None

    def _set_status(self, text: str, color_key: str):
        """Met à jour l'indicateur de statut."""
        self.status_color_key = color_key
        self.status_indicator.setText(text)
        self.status_indicator.setStyleSheet(
            f"color: {self.colors[color_key]}; font-weight: 600; font-size: 12px;"
        )

    def _apply_theme_styles(self):
        """Applique le thème courant à tous les widgets visibles."""
        self.toolbar.setStyleSheet(f"""
            QFrame {{
                background-color: {self.colors['light']};
                border-top-left-radius: 12px;
                border-top-right-radius: 12px;
                border-bottom: 2px solid {self.colors['primary']}20;
                padding: 8px;
            }}
        """)

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

        for btn in self.toolbar_buttons:
            btn.setStyleSheet(f"""
                QPushButton {{
                    background-color: transparent;
                    color: {self.colors['primary']};
                    border: 1px solid {self.colors['primary']}30;
                    border-radius: 6px;
                    padding: 5px 12px;
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

        self.excel_container.setStyleSheet(f"""
            QFrame {{
                background-color: {self.colors['card-bg']};
                border: 2px solid {self.colors['primary']}18;
                border-bottom-left-radius: 12px;
                border-bottom-right-radius: 12px;
            }}
        """)

        self.placeholder.setStyleSheet(f"""
            QFrame {{
                background-color: {self.colors['primary']}08;
                border: 2px dashed {self.colors['primary']}30;
                border-radius: 8px;
            }}
        """)
        self.placeholder_title.setStyleSheet(
            f"color: {self.colors['primary']}; font-size: 20px; font-weight: 600;"
        )
        self.placeholder_desc.setStyleSheet(f"color: {self.colors['gray']}; font-size: 14px;")
        self.placeholder_tagline.setStyleSheet(
            f"color: {self.colors['secondary']}; font-size: 12px; font-style: italic;"
        )

        self.preview_toolbar.setStyleSheet(f"""
            QFrame {{
                background-color: {self.colors['card-bg']};
                border: 1px solid {self.colors['light-gray']};
                border-radius: 8px;
            }}
        """)
        self.sheet_label.setStyleSheet(f"color: {self.colors['primary']}; font-weight: 600;")
        self.preview_info.setStyleSheet(f"color: {self.colors['gray']}; font-size: 12px;")

        self.sheet_selector.setStyleSheet(f"""
            QComboBox {{
                background-color: {self.colors['card-bg']};
                color: {self.colors['dark']};
                border: 1px solid {self.colors['light-gray']};
                border-radius: 6px;
                padding: 4px 8px;
                min-height: 28px;
            }}
            QComboBox::drop-down {{
                border: none;
                width: 24px;
            }}
            QComboBox QAbstractItemView {{
                background-color: {self.colors['card-bg']};
                color: {self.colors['dark']};
                border: 1px solid {self.colors['light-gray']};
                selection-background-color: {self.colors['primary']}35;
                selection-color: {self.colors['dark']};
            }}
        """)

        self.preview_table.setStyleSheet(f"""
            QTableWidget {{
                background-color: {self.colors['card-bg']};
                color: {self.colors['dark']};
                alternate-background-color: {self.colors['light']};
                gridline-color: {self.colors['light-gray']};
                border: 1px solid {self.colors['light-gray']};
                border-radius: 8px;
                selection-background-color: {self.colors['primary']}55;
                selection-color: white;
            }}
            QHeaderView::section {{
                background-color: {self.colors['light']};
                color: {self.colors['primary']};
                border: 1px solid {self.colors['light-gray']};
                border-left: none;
                border-top: none;
                padding: 6px;
                font-weight: 600;
            }}
            QTableCornerButton::section {{
                background-color: {self.colors['light']};
                border: 1px solid {self.colors['light-gray']};
            }}
        """)

    def _set_loaded_mode(self, loaded: bool):
        """Agrandit la zone de prévisualisation quand un fichier est chargé."""
        self.excel_container.setMinimumHeight(620 if loaded else 420)

    def open_in_external_spreadsheet(self, show_message: bool = True) -> bool:
        """Ouvre le fichier dans un tableur externe (Linux/macOS fallback)."""
        if not self.current_file or not self.current_file.exists():
            if show_message:
                QMessageBox.warning(self, "Erreur GULFCAM", "Aucun fichier valide à ouvrir.")
            return False

        file_path = str(self.current_file.absolute())
        try:
            if IS_WINDOWS and hasattr(os, "startfile"):
                os.startfile(file_path)
            elif self.spreadsheet_command:
                self.external_editor_process = subprocess.Popen(self.spreadsheet_command + [file_path])
            elif shutil.which("xdg-open"):
                self.external_editor_process = subprocess.Popen(["xdg-open", file_path])
            elif shutil.which("open"):
                self.external_editor_process = subprocess.Popen(["open", file_path])
            else:
                raise RuntimeError("Aucun tableur détecté. Installez LibreOffice (libreoffice-calc).")

            self._set_status("Tableur externe ouvert", "success")
            if show_message:
                QMessageBox.information(
                    self,
                    "GULFCAM",
                    "Le fichier est ouvert dans le tableur externe.\n"
                    "Utilisez Ctrl+S dans le tableur pour enregistrer.",
                )
            return True
        except Exception as e:
            self._set_status("Ouverture impossible", "danger")
            if show_message:
                QMessageBox.critical(self, "Erreur GULFCAM", f"Impossible d'ouvrir le tableur:\n{str(e)}")
            return False

    def _load_preview_workbook(self, file_path: Path) -> bool:
        """Charge le classeur en mode prévisualisation/édition intégrée."""
        try:
            from openpyxl import load_workbook as ox_load_workbook

            if self.preview_workbook is not None:
                try:
                    self.preview_workbook.close()
                except Exception:
                    pass

            self.preview_workbook = ox_load_workbook(file_path, data_only=False)
            sheet_names = list(self.preview_workbook.sheetnames)
            if not sheet_names:
                raise RuntimeError("Le fichier ne contient aucune feuille.")

            self.sheet_selector.blockSignals(True)
            self.sheet_selector.clear()
            self.sheet_selector.addItems(sheet_names)
            self.sheet_selector.blockSignals(False)
            self.sheet_selector.setCurrentIndex(0)
            self._populate_preview_table(sheet_names[0])

            self.content_stack.setCurrentWidget(self.preview_page)
            self._set_loaded_mode(True)
            self.is_modified = False
            return True
        except Exception as e:
            self.preview_workbook = None
            self.content_stack.setCurrentWidget(self.placeholder)
            self._set_loaded_mode(False)
            QMessageBox.critical(self, "Erreur GULFCAM", f"Prévisualisation impossible:\n{str(e)}")
            return False

    def _on_sheet_changed(self, sheet_name: str):
        if not sheet_name:
            return
        self._populate_preview_table(sheet_name)

    def _populate_preview_table(self, sheet_name: str):
        """Affiche une feuille dans la grille éditable."""
        if self.preview_workbook is None or sheet_name not in self.preview_workbook.sheetnames:
            return

        ws = self.preview_workbook[sheet_name]
        self.preview_sheet_name = sheet_name

        max_row = max(1, min(ws.max_row or 1, self.preview_max_rows))
        max_col = max(1, min(ws.max_column or 1, self.preview_max_cols))

        from openpyxl.utils import get_column_letter

        self._updating_table = True
        self.preview_table.clear()
        self.preview_table.setRowCount(max_row)
        self.preview_table.setColumnCount(max_col)
        self.preview_table.setHorizontalHeaderLabels([get_column_letter(i) for i in range(1, max_col + 1)])
        self.preview_table.setVerticalHeaderLabels([str(i) for i in range(1, max_row + 1)])

        for r in range(1, max_row + 1):
            for c in range(1, max_col + 1):
                value = ws.cell(row=r, column=c).value
                if value is None:
                    continue
                item = QTableWidgetItem(str(value))
                self.preview_table.setItem(r - 1, c - 1, item)

        self._updating_table = False

        clipped = ""
        if ws.max_row > self.preview_max_rows or ws.max_column > self.preview_max_cols:
            clipped = f" (affichage limité à {self.preview_max_rows} x {self.preview_max_cols})"
        self.preview_info.setText(
            f"{sheet_name} • {ws.max_row} lignes x {ws.max_column} colonnes{clipped}"
        )

    def _coerce_cell_value(self, text: str):
        value = text.strip()
        if value == "":
            return None
        if value.startswith("="):
            return value
        try:
            if value.isdigit() or (value.startswith("-") and value[1:].isdigit()):
                return int(value)
            normalized = value.replace(",", ".")
            return float(normalized)
        except Exception:
            return value

    def _on_preview_item_changed(self, item: QTableWidgetItem):
        """Synchronise les modifications de la grille vers le workbook."""
        if self._updating_table or self.preview_workbook is None or not self.preview_sheet_name:
            return

        try:
            ws = self.preview_workbook[self.preview_sheet_name]
            row = item.row() + 1
            col = item.column() + 1
            ws.cell(row=row, column=col).value = self._coerce_cell_value(item.text())
            self.is_modified = True
            self._set_status("Modifications locales", "warning")
        except Exception as e:
            QMessageBox.warning(self, "Erreur GULFCAM", f"Impossible d'écrire la cellule:\n{str(e)}")
        
    def setup_ui(self):
        """Configure l'interface Excel avec couleurs GULFCAM"""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        
        # Barre d'outils Excel
        self.toolbar = QFrame()
        self.toolbar.setMinimumHeight(68)
        toolbar_layout = QHBoxLayout(self.toolbar)
        toolbar_layout.setContentsMargins(12, 10, 12, 10)
        toolbar_layout.setSpacing(10)
        
        # Indicateur de fichier
        self.file_indicator = QLabel("Aucun fichier chargé")
        self.file_indicator.setMinimumHeight(34)
        self.file_indicator.setAlignment(Qt.AlignVCenter | Qt.AlignLeft)
        
        # Statut
        self.status_indicator = QLabel("Prêt")
        self.status_indicator.setStyleSheet(f"""
            QLabel {{
                color: {self.colors['success']};
                font-weight: 600;
                font-size: 12px;
            }}
        """)
        self.status_indicator.setMinimumHeight(30)
        self.status_indicator.setAlignment(Qt.AlignVCenter | Qt.AlignLeft)
        
        # Boutons
        self.save_btn = QPushButton("Enregistrer")
        self.save_btn.setIcon(
            ui_icon(
                self,
                "fa5s.save",
                color=self.colors["primary"],
                fallback=QStyle.StandardPixmap.SP_DialogSaveButton,
            )
        )
        self.save_btn.setCursor(Qt.PointingHandCursor)
        self.save_btn.clicked.connect(self.save_changes)
        
        self.refresh_btn = QPushButton("Recharger")
        self.refresh_btn.setIcon(
            ui_icon(
                self,
                "fa5s.sync-alt",
                color=self.colors["primary"],
                fallback=QStyle.StandardPixmap.SP_BrowserReload,
            )
        )
        self.refresh_btn.setCursor(Qt.PointingHandCursor)
        self.refresh_btn.clicked.connect(self.refresh)
        
        embed_label = "Forcer intégration" if self.use_windows_com else "Ouvrir le tableur"
        self.embed_btn = QPushButton(embed_label)
        embed_fa_icon = "fa5s.object-group" if self.use_windows_com else "fa5s.external-link-alt"
        embed_fallback = (
            QStyle.StandardPixmap.SP_CommandLink if self.use_windows_com else QStyle.StandardPixmap.SP_DialogOpenButton
        )
        self.embed_btn.setIcon(
            ui_icon(
                self,
                embed_fa_icon,
                color=self.colors["primary"],
                fallback=embed_fallback,
            )
        )
        self.embed_btn.setCursor(Qt.PointingHandCursor)
        if self.use_windows_com:
            self.embed_btn.clicked.connect(self.force_embed_excel_now)
        else:
            self.embed_btn.clicked.connect(self.open_in_external_spreadsheet)
        self.toolbar_buttons = [self.save_btn, self.refresh_btn, self.embed_btn]
        for btn in self.toolbar_buttons:
            btn.setMinimumHeight(34)
            btn.setIconSize(QSize(16, 16))
        
        toolbar_layout.addWidget(self.file_indicator)
        toolbar_layout.addSpacing(20)
        toolbar_layout.addWidget(self.status_indicator)
        toolbar_layout.addStretch()
        toolbar_layout.addWidget(self.save_btn)
        toolbar_layout.addWidget(self.refresh_btn)
        toolbar_layout.addWidget(self.embed_btn)
        
        # Conteneur Excel
        self.excel_container = QFrame()
        self.excel_container.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.excel_container.setMinimumHeight(420)
        
        # Layout du conteneur
        container_layout = QVBoxLayout(self.excel_container)
        container_layout.setContentsMargins(10, 10, 10, 10)
        
        # Message placeholder
        self.placeholder = QFrame()
        
        placeholder_layout = QVBoxLayout(self.placeholder)
        placeholder_layout.setAlignment(Qt.AlignCenter)
        placeholder_layout.setSpacing(15)
        
        title_text = "GULFCAM Excel Intégré" if self.use_windows_com else "GULFCAM Tableur Linux"
        self.placeholder_title = QLabel(title_text)
        self.placeholder_title.setAlignment(Qt.AlignCenter)
        
        desc_text = (
            "Déposez votre fichier Excel ici\nou utilisez le bouton \"Charger\""
            if self.use_windows_com
            else "Chargez un fichier .xlsx pour l'éditer ici\nou cliquez sur \"Ouvrir le tableur\""
        )
        self.placeholder_desc = QLabel(desc_text)
        self.placeholder_desc.setAlignment(Qt.AlignCenter)
        
        self.placeholder_tagline = QLabel(GULFCAM_TAGLINE)
        self.placeholder_tagline.setAlignment(Qt.AlignCenter)
        
        placeholder_layout.addWidget(self.placeholder_title)
        placeholder_layout.addWidget(self.placeholder_desc)
        placeholder_layout.addWidget(self.placeholder_tagline)

        # Vue d'édition intégrée (fallback Linux/macOS)
        self.preview_page = QFrame()
        preview_layout = QVBoxLayout(self.preview_page)
        preview_layout.setContentsMargins(6, 6, 6, 6)
        preview_layout.setSpacing(8)

        self.preview_toolbar = QFrame()
        preview_toolbar_layout = QHBoxLayout(self.preview_toolbar)
        preview_toolbar_layout.setContentsMargins(10, 6, 10, 6)
        preview_toolbar_layout.setSpacing(10)

        self.sheet_label = QLabel("Feuille :")
        self.sheet_selector = QComboBox()
        self.sheet_selector.setMinimumWidth(220)
        self.sheet_selector.currentTextChanged.connect(self._on_sheet_changed)

        self.preview_info = QLabel("Prévisualisation locale")

        preview_toolbar_layout.addWidget(self.sheet_label)
        preview_toolbar_layout.addWidget(self.sheet_selector)
        preview_toolbar_layout.addWidget(self.preview_info, 1)

        self.preview_table = QTableWidget()
        self.preview_table.setAlternatingRowColors(True)
        self.preview_table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.preview_table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.preview_table.horizontalHeader().setDefaultSectionSize(130)
        self.preview_table.verticalHeader().setDefaultSectionSize(24)
        self.preview_table.itemChanged.connect(self._on_preview_item_changed)

        preview_layout.addWidget(self.preview_toolbar)
        preview_layout.addWidget(self.preview_table, 1)

        self.content_stack = QStackedWidget()
        self.content_stack.addWidget(self.placeholder)
        self.content_stack.addWidget(self.preview_page)
        self.content_stack.setCurrentWidget(self.placeholder)

        container_layout.addWidget(self.content_stack, 1)
        
        layout.addWidget(self.toolbar)
        layout.addWidget(self.excel_container, 1)
        self._apply_theme_styles()
        self._set_status("Prêt", self.status_color_key)
        
        # Timer pour forcer le reparenting
        self.embed_timer = QTimer()
        if self.use_windows_com:
            self.embed_timer.timeout.connect(self.force_embed_excel)
            self.embed_timer.start(1000)
        
    def force_embed_excel_now(self):
        """Force l'intégration d'Excel immédiatement"""
        if not self.use_windows_com:
            self.open_in_external_spreadsheet()
            return
        self.retry_count = 0
        self.force_embed_excel()
        
    def force_embed_excel(self):
        """Force l'intégration d'Excel dans l'application"""
        if not self.use_windows_com:
            return
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
                
                print(f"Excel intégré avec succès: {excel_hwnd}")
                self.retry_count = 0
                self.embed_timer.stop()
                
            else:
                self.retry_count += 1
                if self.retry_count > self.max_retries:
                    print("Impossible de trouver la fenêtre Excel")
                    self.status_indicator.setText("Échec intégration")
                    self.status_indicator.setStyleSheet(f"color: {self.colors['danger']}; font-weight: 600;")
                    
        except Exception as e:
            print(f"Erreur embedding Excel: {e}")
            self.retry_count += 1
            if self.retry_count > self.max_retries:
                self.embed_timer.stop()
                
    def resize_excel_window(self):
        """Redimensionne Excel pour remplir le conteneur"""
        if not self.use_windows_com:
            return
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
            self.file_indicator.setText(file_path.name)

            if not self.use_windows_com:
                if self._load_preview_workbook(file_path):
                    self._set_status("Prévisualisation chargée", "success")
                    return True
                return False
            
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
            self.content_stack.setCurrentWidget(self.placeholder)
            self._set_loaded_mode(True)
            
            # Forcer l'embedding
            self.embed_timer.start(1000)
            self.force_embed_excel()
            
            self._set_status("Chargé", "success")
            
            return True
            
        except Exception as e:
            QMessageBox.critical(self, "Erreur GULFCAM", f"Impossible de charger le fichier:\n{str(e)}")
            return False
            
    def save_changes(self):
        """Enregistre les modifications"""
        if not self.current_file:
            QMessageBox.warning(self, "Erreur GULFCAM", "Aucun fichier chargé.")
            return

        if not self.use_windows_com:
            if self.preview_workbook is None:
                QMessageBox.warning(self, "Erreur GULFCAM", "Aucune prévisualisation à enregistrer.")
                return
            try:
                self.preview_workbook.save(self.current_file)
                self.is_modified = False
                self._set_status("Enregistré", "success")
                QMessageBox.information(self, "GULFCAM", "Modifications enregistrées avec succès!")
            except Exception as e:
                QMessageBox.critical(self, "Erreur GULFCAM", f"Impossible d'enregistrer:\n{str(e)}")
            return

        if self.workbook:
            try:
                self.workbook.Save()
                self._set_status("Enregistré", "success")
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
                if self.use_windows_com:
                    self.load_file(self.current_file)
                else:
                    if self._load_preview_workbook(self.current_file):
                        self._set_status("Rechargé", "success")
                    else:
                        self._set_status("Rechargement échoué", "danger")
                
    def closeEvent(self, event):
        """Fermeture propre"""
        if self.embed_timer:
            self.embed_timer.stop()
        if self.use_windows_com:
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
        else:
            if self.preview_workbook is not None:
                try:
                    self.preview_workbook.close()
                except Exception:
                    pass
        event.accept()

    def set_colors(self, colors):
        """Met à jour les couleurs et réapplique le thème."""
        self.colors = colors
        self._apply_theme_styles()
        self._set_status(self.status_indicator.text(), self.status_color_key)

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
        header_layout = QGridLayout(header)
        header_layout.setContentsMargins(0, 0, 0, 0)
        header_layout.setHorizontalSpacing(16)
        header_layout.setColumnStretch(0, 1)
        header_layout.setColumnStretch(1, 1)
        header_layout.setColumnStretch(2, 1)
        
        # Logo GULFCAM avec tagline
        logo_container = QFrame()
        logo_layout = QHBoxLayout(logo_container)
        logo_layout.setSpacing(16)

        logo_label = QLabel()
        logo_pixmap = load_logo_pixmap(160, 110)
        if not logo_pixmap.isNull():
            logo_label.setPixmap(logo_pixmap)
            logo_label.setFixedSize(164, 114)
            logo_label.setAlignment(Qt.AlignCenter)
        else:
            logo_label.setText("G")
            logo_label.setStyleSheet(f"""
                font-size: 32px;
                font-weight: 800;
                color: white;
                background-color: {self.colors['primary']};
                border-radius: 10px;
                padding: 8px 16px;
            """)

        brand_text = QVBoxLayout()
        brand_text.setSpacing(2)
        brand_name = QLabel("GULFCAM DSF")
        brand_name.setStyleSheet(f"""
            font-size: 24px;
            font-weight: 800;
            color: {self.colors['primary']};
            letter-spacing: 1px;
        """)
        tagline = QLabel(GULFCAM_TAGLINE)
        tagline.setStyleSheet(f"""
            color: {self.colors['secondary']};
            font-size: 14px;
            font-style: italic;
        """)
        brand_text.addWidget(brand_name)
        brand_text.addWidget(tagline)

        logo_layout.addWidget(logo_label)
        logo_layout.addLayout(brand_text)
        
        date_label = QLabel(time.strftime("%d %B %Y"))
        date_label.setAlignment(Qt.AlignCenter)
        date_label.setStyleSheet(f"""
            color: {self.colors['gray']};
            font-size: 16px;
            font-weight: 600;
            padding: 10px 20px;
            background-color: {self.colors['light-gray']};
            border-radius: 16px;
        """)

        lightgroup_label = QLabel()
        lightgroup_pixmap = load_lightgroup_pixmap(250, 160)
        if not lightgroup_pixmap.isNull():
            lightgroup_label.setPixmap(lightgroup_pixmap)
            lightgroup_label.setFixedSize(154, 64)
            lightgroup_label.setAlignment(Qt.AlignCenter)
        else:
            lightgroup_label.setText("LIGHTGROUP")
            lightgroup_label.setAlignment(Qt.AlignCenter)
            lightgroup_label.setStyleSheet(f"""
                color: {self.colors['primary']};
                font-size: 14px;
                font-weight: 700;
                border: 1px solid {self.colors['light-gray']};
                border-radius: 10px;
                padding: 8px 12px;
            """)

        header_layout.addWidget(logo_container, 0, 0, Qt.AlignLeft | Qt.AlignVCenter)
        header_layout.addWidget(date_label, 0, 1, Qt.AlignCenter)
        header_layout.addWidget(lightgroup_label, 0, 2, Qt.AlignRight | Qt.AlignVCenter)
        
        layout.addWidget(header)
        
        # Message de bienvenue
        welcome_card = ModernCard(colors=self.colors)
        welcome_layout = QHBoxLayout(welcome_card)
        
        welcome_icon = QLabel()
        welcome_icon.setPixmap(
            ui_icon(
                self,
                "fa5s.bolt",
                color=self.colors["secondary"],
                fallback=QStyle.StandardPixmap.SP_ComputerIcon,
            ).pixmap(28, 28)
        )
        welcome_icon.setFixedSize(34, 34)
        welcome_icon.setAlignment(Qt.AlignCenter)
        
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
            ("Nouveau DSF", "fa5s.file-alt", QStyle.StandardPixmap.SP_FileIcon, self.colors['primary']),
            ("Charger balance", "fa5s.file-import", QStyle.StandardPixmap.SP_DirOpenIcon, self.colors['success']),
            ("Rapports", "fa5s.chart-line", QStyle.StandardPixmap.SP_FileDialogDetailedView, self.colors['secondary']),
            ("Paramètres", "fa5s.cogs", QStyle.StandardPixmap.SP_FileDialogContentsView, self.colors['gray']),
        ]
        
        for i, (text, fa_name, icon_sp, color) in enumerate(actions):
            btn = QPushButton(text)
            btn.setIcon(ui_icon(self, fa_name, color=color, fallback=icon_sp))
            btn.setIconSize(QSize(18, 18))
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
        self.setObjectName("mainWorkView")
        self.setStyleSheet(f"""
            QWidget#mainWorkView {{
                background-color: {self.colors['light']};
            }}
        """)

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        # Barre de navigation
        nav_bar = QFrame()
        nav_bar.setStyleSheet(f"""
            QFrame {{
                background-color: {self.colors['card-bg']};
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

        logo_image = QLabel()
        logo_pixmap = load_logo_pixmap(80, 60)
        if not logo_pixmap.isNull():
            logo_image.setPixmap(logo_pixmap)
            logo_image.setAlignment(Qt.AlignCenter)
            logo_image.setFixedSize(84, 64)
        else:
            logo_image.setText("G")
            logo_image.setStyleSheet(f"""
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
        
        logo_layout.addWidget(logo_image)
        logo_layout.addWidget(logo_full)
        
        # Navigation
        nav_buttons = [
            ("Accueil", 0, "fa5s.home", QStyle.StandardPixmap.SP_DirHomeIcon),
            ("Balance", 1, "fa5s.balance-scale", QStyle.StandardPixmap.SP_DriveHDIcon),
            ("DSF", 2, "fa5s.file-signature", QStyle.StandardPixmap.SP_FileIcon),
            ("Rapports", 3, "fa5s.chart-bar", QStyle.StandardPixmap.SP_FileDialogDetailedView),
        ]
        
        nav_layout.addWidget(logo_container)
        nav_layout.addSpacing(40)
        
        for text, idx, fa_name, icon_sp in nav_buttons:
            btn = QPushButton(text)
            btn.setIcon(ui_icon(self, fa_name, color=self.colors["gray"], fallback=icon_sp))
            btn.setIconSize(QSize(16, 16))
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
        theme_btn = QPushButton("Thème")
        theme_btn.setIcon(
            ui_icon(
                self,
                "fa5s.adjust",
                color=self.colors["primary"],
                fallback=QStyle.StandardPixmap.SP_DesktopIcon,
            )
        )
        theme_btn.setIconSize(QSize(16, 16))
        theme_btn.setCursor(Qt.PointingHandCursor)
        theme_btn.setMinimumHeight(40)
        theme_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {self.colors['primary']}10;
                color: {self.colors['primary']};
                border: none;
                border-radius: 20px;
                font-size: 13px;
                font-weight: 600;
                padding: 0 12px;
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
        scroll.setStyleSheet(f"""
            QScrollArea {{
                border: none;
                background-color: {self.colors['light']};
            }}
        """)
        
        content = QWidget()
        content.setObjectName("scrollContent")
        content.setStyleSheet(f"""
            QWidget#scrollContent {{
                background-color: {self.colors['light']};
            }}
        """)
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
            self.controller.statusBar().showMessage(f"GULFCAM - Balance chargée: {Path(path).name}")
            
    def generate_dsf(self):
        if not self.balance_file:
            QMessageBox.warning(self, "Erreur GULFCAM", "Veuillez d'abord charger une balance")
            return
            
        QMessageBox.information(self, "GULFCAM", "Fonctionnalité de génération DSF à implémenter")
        
    def export_dsf(self):
        if not self.excel_widget.current_file or not self.excel_widget.current_file.exists():
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
        """Crée l'icône de l'application depuis OIP.png."""
        logo_pixmap = load_logo_pixmap(128, 128)
        if not logo_pixmap.isNull():
            return QIcon(logo_pixmap)

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
        status_bar.showMessage(f"{GULFCAM_LOGO_TEXT} • {GULFCAM_TAGLINE} • Application prête")
        
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
        
        self.statusBar().showMessage(f"Mode {'sombre' if self.settings.dark_mode else 'clair'} GULFCAM activé")
        
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
    com_initialized = False
    if WINDOWS_COM_AVAILABLE and pythoncom is not None:
        # Initialiser COM uniquement sur Windows avec pywin32 disponible
        pythoncom.CoInitialize()
        com_initialized = True
    
    app = QApplication([])
    
    # Police par défaut
    font = QFont("Segoe UI", 10)
    app.setFont(font)
    
    # Variable pour stocker la fenêtre principale
    main_window = None
    startup_state = {"window": None}

    def load_settings_step():
        AppSettings.load()

    def load_assets_step():
        load_logo_pixmap(185, 195)
        load_lightgroup_pixmap(150, 60)

    def build_ui_step():
        startup_state["window"] = LightDSFMapperWindow()

    startup_tasks = [
        ("Chargement des paramètres", load_settings_step),
        ("Chargement des ressources graphiques", load_assets_step),
        ("Initialisation de l'interface", build_ui_step),
    ]
    
    def show_main_window():
        nonlocal main_window
        main_window = startup_state.get("window")
        if main_window is None:
            main_window = LightDSFMapperWindow()
        main_window.show()
    
    # Afficher le splash screen
    splash = SplashScreen(show_main_window, startup_tasks=startup_tasks, min_duration_ms=5000)
    splash.show()
    
    result = app.exec()
    
    if com_initialized:
        pythoncom.CoUninitialize()
    
    return result

if __name__ == "__main__":
    main()
