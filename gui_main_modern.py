#!/usr/bin/env python3
"""
Modern GUI for Excel Comparison Tool - GridKit
Enhanced visual design while preserving all functionality
Built with PySide6
"""

import sys
import os

# Version check
if sys.version_info < (3, 8):
    print("❌ Error: Python 3.8 or higher is required")
    print(f"   Current version: {sys.version}")
    print("   Please upgrade Python from https://python.org")
    sys.exit(1)

from pathlib import Path
from datetime import datetime
import time
import platform
import pandas as pd

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QGroupBox, QCheckBox,
    QProgressBar, QMessageBox, QScrollArea, QGridLayout, QLineEdit,
    QComboBox, QInputDialog, QFrame, QSizePolicy
)
from PySide6.QtCore import Qt, QThread, Signal, QSettings
from PySide6.QtGui import QFont, QAction, QKeySequence, QDragEnterEvent, QDropEvent, QIcon

from src.core import ComparisonEngine, ComparisonConfig, AlignmentMethod
from src.reports.report_generator import generate_comparison_report


# =========================
# Worker Thread
# =========================
class ComparisonWorker(QThread):
    progress = Signal(str)
    finished = Signal(object)
    error = Signal(str)

    def __init__(self, df_a, df_b, config, file_a_path, file_b_path):
        super().__init__()
        self.df_a = df_a
        self.df_b = df_b
        self.config = config
        self.file_a_path = file_a_path
        self.file_b_path = file_b_path

    def run(self):
        try:
            self.progress.emit("🔍 Comparing files...")
            engine = ComparisonEngine(self.config)
            result = engine.compare(self.df_a, self.df_b)

            self.progress.emit("📄 Generating Excel report...")
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = f"comparison_report_{timestamp}.xlsx"

            generate_comparison_report(
                output_path=output_file,
                summary=result.summary,
                aligned_data=result.aligned_data,
                metadata=result.comparison_metadata,
                file_a_path=self.file_a_path,
                file_b_path=self.file_b_path
            )

            self.finished.emit({
                "result": result,
                "output_path": Path(output_file).resolve()
            })

        except Exception as e:
            self.error.emit(str(e))


# =========================
# Main GUI - Modernized
# =========================
class ExcelComparisonGUI(QMainWindow):
    # Modern color scheme
    COLOR_PRIMARY = "#667eea"          # Purple accent
    COLOR_PRIMARY_DARK = "#5568d3"     # Darker purple for hover
    COLOR_SUCCESS = "#48bb78"          # Green
    COLOR_ERROR = "#e53e3e"            # Red
    COLOR_WARNING = "#ed8936"          # Orange
    COLOR_TEXT_PRIMARY = "#1a202c"     # Almost black
    COLOR_TEXT_SECONDARY = "#718096"   # Gray
    COLOR_TEXT_TERTIARY = "#a0aec0"    # Light gray
    COLOR_BG_LIGHT = "#f7fafc"         # Very light gray
    COLOR_BG_WHITE = "#ffffff"         # White
    COLOR_BORDER = "#e2e8f0"           # Light border

    def __init__(self):
        super().__init__()
        self.file_a_path = None
        self.file_b_path = None
        self.file_a_sheet = None
        self.file_b_sheet = None
        self.df_a = None
        self.df_b = None
        self.key_checkboxes = []
        self.worker = None
        self.start_time = None
       
        # Settings
        self.settings = QSettings("ExcelCompTool", "ExcelComparisonTool")
        self.last_directory = self.settings.value("last_directory", str(Path.home()))
       
        self.init_ui()
        self.load_settings()
       
        # Enable drag and drop
        self.setAcceptDrops(True)

    # ---------- UI ----------
    def init_ui(self):
        self.setWindowTitle("GridKit – Excel Comparison Tool v1.0")
        self.setMinimumSize(900, 650)
        self.resize(1000, 800)

        # Modern gradient background
        self.setStyleSheet(f"""
            QMainWindow {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 {self.COLOR_BG_LIGHT}, stop:1 #edf2f7);
            }}
        """)

        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setSpacing(20)
        main_layout.setContentsMargins(24, 24, 24, 24)

        # Modern header
        title = QLabel("GridKit")
        title.setFont(self.ui_font(size=32, bold=True))
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet(f"""
            color: {self.COLOR_TEXT_PRIMARY};
            padding: 10px;
        """)
        main_layout.addWidget(title)

        subtitle = QLabel("Compare two Excel files and highlight differences")
        subtitle.setAlignment(Qt.AlignCenter)
        subtitle.setStyleSheet(f"""
            color: {self.COLOR_TEXT_SECONDARY};
            font-size: 14pt;
            padding-bottom: 10px;
        """)
        main_layout.addWidget(subtitle)

        # Scrollable content area
        content_scroll = QScrollArea()
        content_scroll.setWidgetResizable(True)
        content_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        content_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        content_scroll.setFrameShape(QFrame.Shape.NoFrame)
        content_scroll.setStyleSheet("QScrollArea { border: none; background: transparent; }")
        
        content_widget = QWidget()
        content_layout = QVBoxLayout(content_widget)
        content_layout.setSpacing(16)
        content_layout.setContentsMargins(0, 0, 0, 0)
        
        # Sections in scrollable area
        content_layout.addWidget(self.create_file_section())
        content_layout.addWidget(self.create_config_section())
        content_layout.addStretch()
        
        content_scroll.setWidget(content_widget)
        main_layout.addWidget(content_scroll, stretch=1)
        
        # Compare section
        main_layout.addWidget(self.create_compare_section())

        self.statusBar().showMessage("Ready – drag & drop Excel files or use the Browse buttons")
        self.statusBar().setStyleSheet(f"color: {self.COLOR_TEXT_SECONDARY}; font-size: 11pt;")
       
        # Keyboard shortcuts
        self.setup_shortcuts()
        
        # Connect signals
        self.tiebreaker_combo.currentIndexChanged.connect(self.on_tiebreaker_changed)

    def ui_font(self, size=10, bold=False):
        font = QFont("Segoe UI", size)
        if bold:
            font.setWeight(QFont.Weight.Bold)
        return font

    def setup_shortcuts(self):
        """Setup keyboard shortcuts"""
        compare_action = QAction("Compare", self)
        compare_action.setShortcut(QKeySequence("Ctrl+Return"))
        compare_action.triggered.connect(self.run_comparison)
        self.addAction(compare_action)
        self.compare_btn.setToolTip("Click or press Ctrl+Enter to compare")

    # ---------- Modern Card Style ----------
    def card_style(self):
        return f"""
            QGroupBox {{
                background: {self.COLOR_BG_WHITE};
                border-radius: 12px;
                padding: 14px;
                margin-top: 12px;
                border: 1px solid {self.COLOR_BORDER};
            }}
            QGroupBox::title {{
                color: {self.COLOR_TEXT_PRIMARY};
                font-size: 16pt;
                font-weight: bold;
                padding: 0 8px;
                subcontrol-origin: margin;
                subcontrol-position: top left;
                left: 12px;
                top: -10px;
                background: {self.COLOR_BG_WHITE};
            }}
        """

    # ---------- File Section ----------
    def create_file_section(self):
        group = QGroupBox("📁 1. Select Files")
        group.setStyleSheet(self.card_style())
        layout = QVBoxLayout(group)
        layout.setSpacing(12)
        layout.setContentsMargins(16, 28, 16, 16)

        # File A section
        file_a_layout = QGridLayout()
        file_a_layout.setSpacing(8)
        
        lbl_a = QLabel("File A:")
        lbl_a.setStyleSheet(f"font-size: 12pt; font-weight: 600; color: {self.COLOR_TEXT_PRIMARY};")
       
        self.file_a_display = QLineEdit()
        self.file_a_display.setPlaceholderText("Drag & drop, browse, or paste file path...")
        self.file_a_display.setStyleSheet(f"""
            QLineEdit {{
                padding: 10px 12px;
                font-size: 11pt;
                background: #f8f9fa;
                color: {self.COLOR_TEXT_PRIMARY};
                border: 2px solid {self.COLOR_BORDER};
                border-radius: 8px;
            }}
            QLineEdit:focus {{
                border-color: {self.COLOR_PRIMARY};
                background: white;
            }}
        """)
        self.file_a_display.textChanged.connect(lambda: self.on_file_path_changed("A"))
       
        btn_a = QPushButton("Browse...")
        btn_a.setFixedWidth(100)
        btn_a.setStyleSheet(self.secondary_button_style())
        btn_a.clicked.connect(lambda: self.select_file("A"))

        file_a_layout.addWidget(lbl_a, 0, 0)
        file_a_layout.addWidget(self.file_a_display, 0, 1)
        file_a_layout.addWidget(btn_a, 0, 2)
        file_a_layout.setColumnStretch(1, 1)
        
        tip_a = QLabel("💡 Tip: Put your original (before) file here")
        tip_a.setStyleSheet(f"font-size: 10pt; color: {self.COLOR_TEXT_SECONDARY}; padding-left: 8px; font-style: italic;")
        file_a_layout.addWidget(tip_a, 1, 1)
        
        layout.addLayout(file_a_layout)

        # File B section
        file_b_layout = QGridLayout()
        file_b_layout.setSpacing(8)
        
        lbl_b = QLabel("File B:")
        lbl_b.setStyleSheet(f"font-size: 12pt; font-weight: 600; color: {self.COLOR_TEXT_PRIMARY};")
       
        self.file_b_display = QLineEdit()
        self.file_b_display.setPlaceholderText("Drag & drop, browse, or paste file path...")
        self.file_b_display.setStyleSheet(f"""
            QLineEdit {{
                padding: 10px 12px;
                font-size: 11pt;
                background: #f0f8ff;
                color: {self.COLOR_TEXT_PRIMARY};
                border: 2px solid {self.COLOR_BORDER};
                border-radius: 8px;
            }}
            QLineEdit:focus {{
                border-color: {self.COLOR_PRIMARY};
                background: white;
            }}
        """)
        self.file_b_display.textChanged.connect(lambda: self.on_file_path_changed("B"))
       
        btn_b = QPushButton("Browse...")
        btn_b.setFixedWidth(100)
        btn_b.setStyleSheet(self.secondary_button_style())
        btn_b.clicked.connect(lambda: self.select_file("B"))

        file_b_layout.addWidget(lbl_b, 0, 0)
        file_b_layout.addWidget(self.file_b_display, 0, 1)
        file_b_layout.addWidget(btn_b, 0, 2)
        file_b_layout.setColumnStretch(1, 1)
        
        tip_b = QLabel("💡 Tip: Put your updated (after) file here to see what changed")
        tip_b.setStyleSheet(f"font-size: 10pt; color: {self.COLOR_TEXT_SECONDARY}; padding-left: 8px; font-style: italic;")
        file_b_layout.addWidget(tip_b, 1, 1)
        
        layout.addLayout(file_b_layout)

        return group

    def on_file_path_changed(self, which):
        """Handle manual file path entry"""
        if which == "A":
            path = self.file_a_display.text().strip()
        else:
            path = self.file_b_display.text().strip()
       
        if not path:
            self.clear_file(which)
            self.update_compare_button_state()
            return
        
        valid_extensions = ('.xlsx', '.xls', '.xlsm')
        if not path.lower().endswith(valid_extensions):
            QMessageBox.warning(self, "Invalid File Type", 
                "Please enter a valid Excel file.\n\nSupported formats: .xlsx, .xls, or .xlsm")
            self.clear_file(which)
            self.update_compare_button_state()
            return
        
        path_obj = Path(path)
        if not path_obj.exists():
            QMessageBox.warning(self, "File Not Found",
                "The file path you entered does not exist.")
            self.clear_file(which)
            self.update_compare_button_state()
            return
        
        if not path_obj.is_file():
            QMessageBox.warning(self, "Invalid Path",
                "The path you entered is not a file.")
            self.clear_file(which)
            self.update_compare_button_state()
            return
        
        self.load_file_path(path, which)

    # ---------- Config Section ----------
    def create_config_section(self):
        self.config_group = QGroupBox("⚙️ 2. Configure Comparison")
        self.config_group.setEnabled(False)
        self.config_group.setStyleSheet(self.card_style())
        layout = QVBoxLayout(self.config_group)
        layout.setSpacing(14)
        layout.setContentsMargins(16, 28, 16, 16)

        # Description
        desc_label = QLabel("Use key-based when rows are identified by IDs. Use position-based when rows line up by row number.")
        desc_label.setStyleSheet(f"font-size: 11pt; color: {self.COLOR_TEXT_SECONDARY}; padding-bottom: 8px;")
        desc_label.setWordWrap(True)
        layout.addWidget(desc_label)

        # Mode selection with modern radio-style checkboxes
        mode_container = QFrame()
        mode_container.setStyleSheet(f"""
            QFrame {{
                background: {self.COLOR_BG_LIGHT};
                border: 1px solid {self.COLOR_BORDER};
                border-radius: 8px;
                padding: 12px;
            }}
        """)
        mode_layout = QHBoxLayout(mode_container)
        mode_layout.setSpacing(16)
       
        self.mode_key_based = QCheckBox("🔑 Key-Based (Row Matching)")
        self.mode_key_based.setChecked(True)
        self.mode_key_based.setStyleSheet(self.modern_checkbox_style())
        self.mode_key_based.toggled.connect(self.on_mode_changed)
       
        self.mode_position_based = QCheckBox("📍 Position-Based (Row 1 → Row 1)")
        self.mode_position_based.setStyleSheet(self.modern_checkbox_style())
        self.mode_position_based.toggled.connect(self.on_mode_changed)
       
        mode_layout.addWidget(self.mode_key_based)
        mode_layout.addWidget(self.mode_position_based)
        mode_layout.addStretch()
       
        layout.addWidget(mode_container)

        # Key columns section
        key_frame = QFrame()
        key_frame.setStyleSheet(f"""
            QFrame {{
                background: {self.COLOR_BG_LIGHT};
                border: 1px solid {self.COLOR_BORDER};
                border-radius: 8px;
                padding: 12px;
            }}
        """)
        key_frame_layout = QVBoxLayout(key_frame)
        key_frame_layout.setSpacing(10)
        
        key_title = QLabel("🔑 Key Columns")
        key_title.setStyleSheet(f"font-size: 12pt; font-weight: 600; color: {self.COLOR_TEXT_PRIMARY};")
        key_frame_layout.addWidget(key_title)
        
        key_subtitle = QLabel("Choose one or more columns that uniquely identify each row (e.g., Policy #)")
        key_subtitle.setStyleSheet(f"font-size: 10pt; color: {self.COLOR_TEXT_SECONDARY};")
        key_subtitle.setWordWrap(True)
        key_frame_layout.addWidget(key_subtitle)
        
        self.key_section = QWidget()
        key_section_layout = QVBoxLayout(self.key_section)
        key_section_layout.setSpacing(8)
        key_section_layout.setContentsMargins(0, 0, 0, 0)
       
        # Placeholder
        self.key_placeholder = QLabel("📋 Load files to see available columns")
        self.key_placeholder.setStyleSheet(f"""
            font-size: 11pt;
            color: {self.COLOR_TEXT_TERTIARY};
            font-style: italic;
            padding: 30px;
            background: white;
            border: 2px dashed {self.COLOR_BORDER};
            border-radius: 8px;
        """)
        self.key_placeholder.setWordWrap(True)
        self.key_placeholder.setAlignment(Qt.AlignCenter)
        key_section_layout.addWidget(self.key_placeholder)

        # Control buttons
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(8)
       
        self.select_all_btn = QPushButton("✓ Select All")
        self.select_all_btn.setStyleSheet(self.tertiary_button_style())
        self.select_all_btn.clicked.connect(lambda: self.toggle_all_keys(True))
        self.select_all_btn.setVisible(False)
       
        self.deselect_all_btn = QPushButton("✗ Deselect All")
        self.deselect_all_btn.setStyleSheet(self.tertiary_button_style())
        self.deselect_all_btn.clicked.connect(lambda: self.toggle_all_keys(False))
        self.deselect_all_btn.setVisible(False)
       
        btn_layout.addWidget(self.select_all_btn)
        btn_layout.addWidget(self.deselect_all_btn)
        btn_layout.addStretch()
        key_section_layout.addLayout(btn_layout)

        # Filter
        self.key_filter = QLineEdit()
        self.key_filter.setPlaceholderText("🔍 Filter columns...")
        self.key_filter.setStyleSheet(f"""
            QLineEdit {{
                padding: 8px 12px;
                font-size: 11pt;
                border: 2px solid {self.COLOR_BORDER};
                border-radius: 6px;
                background: white;
            }}
            QLineEdit:focus {{
                border-color: {self.COLOR_PRIMARY};
            }}
        """)
        self.key_filter.textChanged.connect(self.filter_key_columns)
        self.key_filter.setVisible(False)
        key_section_layout.addWidget(self.key_filter)

        # Scroll area for checkboxes
        self.key_scroll = QScrollArea()
        self.key_scroll.setWidgetResizable(True)
        self.key_scroll.setMaximumHeight(220)
        self.key_scroll.setMinimumHeight(150)
        self.key_scroll.setStyleSheet(f"""
            QScrollArea {{
                border: 2px solid {self.COLOR_BORDER};
                border-radius: 6px;
                background: white;
            }}
        """)
        self.key_scroll.setVisible(False)

        self.key_container = QWidget()
        self.key_container.setStyleSheet("background: white;")
        self.key_grid = QGridLayout(self.key_container)
        self.key_grid.setSpacing(8)
        self.key_grid.setContentsMargins(10, 10, 10, 10)
        self.key_container.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Minimum)

        self.key_scroll.setWidget(self.key_container)
        key_section_layout.addWidget(self.key_scroll)

        # Key count
        self.key_count_label = QLabel("")
        self.key_count_label.setStyleSheet(f"font-size: 10pt; color: {self.COLOR_TEXT_SECONDARY}; padding: 4px;")
        self.key_count_label.setVisible(False)
        key_section_layout.addWidget(self.key_count_label)
       
        key_frame_layout.addWidget(self.key_section)
        layout.addWidget(key_frame)

        # Position-based info
        self.position_info = QLabel(
            "ℹ️ Position-based mode compares files row-by-row without keys.\n"
            "Row 1 in File A is compared to Row 1 in File B, etc."
        )
        self.position_info.setStyleSheet(f"""
            font-size: 11pt;
            color: {self.COLOR_TEXT_SECONDARY};
            background: #e6f7ff;
            border: 1px solid #91d5ff;
            border-radius: 6px;
            padding: 12px;
        """)
        self.position_info.setWordWrap(True)
        self.position_info.setVisible(False)
        layout.addWidget(self.position_info)

        # Advanced options
        self.advanced_expanded = False
        self.advanced_toggle = QPushButton("▼ Advanced Options")
        self.advanced_toggle.setStyleSheet(f"""
            QPushButton {{
                text-align: left;
                padding: 8px;
                font-size: 11pt;
                font-weight: 500;
                background: transparent;
                border: none;
                color: {self.COLOR_PRIMARY};
            }}
            QPushButton:hover {{
                background: {self.COLOR_BG_LIGHT};
                border-radius: 6px;
            }}
        """)
        self.advanced_toggle.clicked.connect(self.toggle_advanced_options)
        layout.addWidget(self.advanced_toggle)
        
        self.advanced_container = QWidget()
        self.advanced_container.setVisible(False)
        advanced_layout = QVBoxLayout(self.advanced_container)
        advanced_layout.setSpacing(10)
        advanced_layout.setContentsMargins(0, 0, 0, 0)
        
        options_layout = QGridLayout()
        options_layout.setSpacing(10)
        options_layout.setColumnStretch(1, 1)
       
        self.tiebreaker_label = QLabel("Tiebreaker Column:")
        self.tiebreaker_label.setStyleSheet(f"font-size: 11pt; font-weight: 500; color: {self.COLOR_TEXT_PRIMARY};")
       
        self.tiebreaker_combo = QComboBox()
        self.tiebreaker_combo.setStyleSheet(f"""
            QComboBox {{
                padding: 8px 12px;
                font-size: 11pt;
                border: 2px solid {self.COLOR_BORDER};
                border-radius: 6px;
                background: white;
            }}
            QComboBox:hover {{
                border-color: {self.COLOR_PRIMARY};
            }}
        """)
        
        options_layout.addWidget(self.tiebreaker_label, 0, 0, Qt.AlignmentFlag.AlignRight)
        options_layout.addWidget(self.tiebreaker_combo, 0, 1)
       
        self.tiebreaker_tip = QLabel("💡 Tip: Use \"Sort by\" when files have same keys but rows are in different order")
        self.tiebreaker_tip.setStyleSheet(f"font-size: 10pt; color: {self.COLOR_TEXT_SECONDARY}; font-style: italic;")
        self.tiebreaker_tip.setVisible(False)
        self.tiebreaker_tip.setWordWrap(True)
        options_layout.addWidget(self.tiebreaker_tip, 1, 0, 1, 2)
       
        self.case_sensitive = QCheckBox("Case Sensitive")
        self.case_sensitive.setStyleSheet(self.modern_checkbox_style())
       
        self.trim_whitespace = QCheckBox("Trim Whitespace")
        self.trim_whitespace.setChecked(True)
        self.trim_whitespace.setStyleSheet(self.modern_checkbox_style())

        options_layout.addWidget(self.case_sensitive, 2, 1)
        options_layout.addWidget(self.trim_whitespace, 3, 1)
        
        advanced_layout.addLayout(options_layout)
        layout.addWidget(self.advanced_container)

        return self.config_group

    def modern_checkbox_style(self):
        return f"""
            QCheckBox {{
                font-size: 11pt;
                color: {self.COLOR_TEXT_PRIMARY};
                spacing: 8px;
            }}

            QCheckBox::indicator {{
                width: 18px;
                height: 18px;
                border-radius: 4px;
                border: 2px solid {self.COLOR_BORDER};
                background: white;
            }}

            QCheckBox::indicator:hover {{
                border-color: {self.COLOR_PRIMARY};
            }}

            QCheckBox::indicator:checked {{
                background-color: {self.COLOR_PRIMARY};
                border-color: {self.COLOR_PRIMARY};
            }}

            QCheckBox::indicator:checked:hover {{
                background-color: {self.COLOR_PRIMARY_DARK};
                border-color: {self.COLOR_PRIMARY_DARK};
            }}
        """




    def toggle_advanced_options(self):
        """Toggle advanced options visibility"""
        self.advanced_expanded = not self.advanced_expanded
        self.advanced_container.setVisible(self.advanced_expanded)
        self.advanced_toggle.setText("▲ Advanced Options" if self.advanced_expanded else "▼ Advanced Options")
        
        if self.advanced_expanded and self.mode_key_based.isChecked():
            self.tiebreaker_label.setVisible(True)
            self.tiebreaker_combo.setVisible(True)
            tiebreaker = self.tiebreaker_combo.currentData()
            self.tiebreaker_tip.setVisible(tiebreaker is not None)
        else:
            if self.advanced_expanded:
                self.tiebreaker_label.setVisible(False)
                self.tiebreaker_combo.setVisible(False)
                self.tiebreaker_tip.setVisible(False)
    
    def on_mode_changed(self):
        """Handle mode change with radio button behavior"""
        sender = self.sender()
        
        if sender == self.mode_key_based and self.mode_key_based.isChecked():
            self.mode_position_based.blockSignals(True)
            self.mode_position_based.setChecked(False)
            self.mode_position_based.blockSignals(False)
            self.key_section.setVisible(True)
            self.position_info.setVisible(False)
            if self.advanced_expanded:
                self.tiebreaker_label.setVisible(True)
                self.tiebreaker_combo.setVisible(True)
                tiebreaker = self.tiebreaker_combo.currentData()
                self.tiebreaker_tip.setVisible(tiebreaker is not None)
            
        elif sender == self.mode_position_based and self.mode_position_based.isChecked():
            self.mode_key_based.blockSignals(True)
            self.mode_key_based.setChecked(False)
            self.mode_key_based.blockSignals(False)
            self.key_section.setVisible(False)
            self.position_info.setVisible(True)
            if self.advanced_expanded:
                self.tiebreaker_label.setVisible(False)
                self.tiebreaker_combo.setVisible(False)
                self.tiebreaker_tip.setVisible(False)
            
        elif not self.mode_key_based.isChecked() and not self.mode_position_based.isChecked():
            if sender == self.mode_key_based:
                self.mode_key_based.blockSignals(True)
                self.mode_key_based.setChecked(True)
                self.mode_key_based.blockSignals(False)
            else:
                self.mode_position_based.blockSignals(True)
                self.mode_position_based.setChecked(True)
                self.mode_position_based.blockSignals(False)

    def on_tiebreaker_changed(self):
        """Handle tiebreaker selection"""
        tiebreaker = self.tiebreaker_combo.currentData()
        self.tiebreaker_tip.setVisible(tiebreaker is not None)

    # ---------- Compare Section ----------
    # ---------- Compare Section ----------
    def create_compare_section(self):
        frame = QFrame()
        frame.setStyleSheet(f"""
            QFrame {{
                background: {self.COLOR_BG_WHITE};
                border-radius: 12px;
                border: 1px solid {self.COLOR_BORDER};
                padding: 16px;
            }}
        """)
        layout = QVBoxLayout(frame)
        layout.setSpacing(12)

        # Status
        self.status_label = QLabel("Ready to compare")
        self.status_label.setStyleSheet(f"font-size: 11pt; color: {self.COLOR_TEXT_SECONDARY};")
        layout.addWidget(self.status_label)

        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0)  # Indeterminate
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        # Buttons
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        self.compare_btn = QPushButton("▶ Compare")
        self.compare_btn.setEnabled(False)
        self.compare_btn.setStyleSheet(self.primary_button_style())
        self.compare_btn.clicked.connect(self.run_comparison)

        self.cancel_btn = QPushButton("✖ Cancel")
        self.cancel_btn.setVisible(False)
        self.cancel_btn.setStyleSheet(self.secondary_button_style())
        self.cancel_btn.clicked.connect(self.cancel_comparison)

        btn_layout.addWidget(self.cancel_btn)
        btn_layout.addWidget(self.compare_btn)
        layout.addLayout(btn_layout)

        return frame
    def run_comparison(self):
        if not self.df_a or not self.df_b:
            QMessageBox.warning(self, "Missing Files", "Please load both files.")
            return

        keys = [cb.text() for cb in self.key_checkboxes if cb.isChecked()]
        if self.mode_key_based.isChecked() and not keys:
            QMessageBox.warning(self, "Missing Keys", "Select at least one key column.")
            return

        config = ComparisonConfig(
            key_columns=keys if self.mode_key_based.isChecked() else None,
            alignment_method=AlignmentMethod.KEY if self.mode_key_based.isChecked()
            else AlignmentMethod.POSITION,
            case_sensitive=self.case_sensitive.isChecked(),
            trim_whitespace=self.trim_whitespace.isChecked(),
            tiebreaker_column=self.tiebreaker_combo.currentData()
        )

        self.worker = ComparisonWorker(
            self.df_a,
            self.df_b,
            config,
            self.file_a_path,
            self.file_b_path
        )

        self.worker.progress.connect(self.on_worker_progress)
        self.worker.finished.connect(self.on_worker_finished)
        self.worker.error.connect(self.on_worker_error)

        self.progress_bar.setVisible(True)
        self.cancel_btn.setVisible(True)
        self.compare_btn.setEnabled(False)
        self.status_label.setText("Starting comparison…")

        self.worker.start()
    def on_worker_progress(self, msg):
        self.status_label.setText(msg)

    def on_worker_finished(self, payload):
        self.progress_bar.setVisible(False)
        self.cancel_btn.setVisible(False)
        self.compare_btn.setEnabled(True)

        output = payload["output_path"]
        self.status_label.setText("Comparison complete ✅")

        QMessageBox.information(
            self,
            "Done",
            f"Report generated:\n\n{output}"
        )

    def on_worker_error(self, error):
        self.progress_bar.setVisible(False)
        self.cancel_btn.setVisible(False)
        self.compare_btn.setEnabled(True)

        QMessageBox.critical(self, "Error", error)
    def cancel_comparison(self):
        if self.worker and self.worker.isRunning():
            self.worker.terminate()
            self.worker.wait()

        self.progress_bar.setVisible(False)
        self.cancel_btn.setVisible(False)
        self.compare_btn.setEnabled(True)
        self.status_label.setText("Comparison cancelled")
    def select_file(self, which):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Excel File",
            self.last_directory,
            "Excel Files (*.xlsx *.xls *.xlsm)"
        )
        if path:
            self.load_file_path(path, which)

    def load_file_path(self, path, which):
        df = pd.read_excel(path)
        if which == "A":
            self.file_a_path = path
            self.df_a = df
            self.file_a_display.setText(path)
        else:
            self.file_b_path = path
            self.df_b = df
            self.file_b_display.setText(path)

        self.last_directory = str(Path(path).parent)
        self.settings.setValue("last_directory", self.last_directory)

        if self.df_a is not None and self.df_b is not None:
            self.populate_key_columns()
            self.config_group.setEnabled(True)
            self.compare_btn.setEnabled(True)
    def populate_key_columns(self):
        for cb in self.key_checkboxes:
            cb.deleteLater()

        self.key_checkboxes.clear()
        self.key_grid.setRowStretch(0, 0)

        columns = list(self.df_a.columns)
        self.key_placeholder.setVisible(False)
        self.key_scroll.setVisible(True)
        self.key_filter.setVisible(True)
        self.select_all_btn.setVisible(True)
        self.deselect_all_btn.setVisible(True)

        for i, col in enumerate(columns):
            cb = QCheckBox(col)
            cb.setStyleSheet(self.modern_checkbox_style())
            self.key_checkboxes.append(cb)
            self.key_grid.addWidget(cb, i // 2, i % 2)

        self.key_count_label.setText(f"{len(columns)} columns available")
        self.key_count_label.setVisible(True)

    # ---------- Drag & Drop ----------
    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        urls = event.mimeData().urls()
        if not urls:
            return

        paths = [url.toLocalFile() for url in urls if url.isLocalFile()]
        excel_files = [p for p in paths if p.lower().endswith((".xlsx", ".xls", ".xlsm"))]

        if not excel_files:
            QMessageBox.warning(
                self,
                "Invalid Drop",
                "Please drop valid Excel files (.xlsx, .xls, .xlsm)."
            )
            return

        if len(excel_files) >= 1:
            self.file_a_display.setText(excel_files[0])
        if len(excel_files) >= 2:
            self.file_b_display.setText(excel_files[1])

    # ---------- Compare Section ----------
    def create_compare_section(self):
        container = QFrame()
        container.setStyleSheet("background: transparent;")
        layout = QVBoxLayout(container)
        layout.setSpacing(12)

        # Progress label
        self.progress_label = QLabel("")
        self.progress_label.setAlignment(Qt.AlignCenter)
        self.progress_label.setStyleSheet(
            f"font-size: 11pt; color: {self.COLOR_TEXT_SECONDARY};"
        )
        layout.addWidget(self.progress_label)

        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setRange(0, 0)  # indeterminate
        layout.addWidget(self.progress_bar)

        # Buttons
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        self.compare_btn = QPushButton("▶ Compare")
        self.compare_btn.setFixedHeight(44)
        self.compare_btn.setEnabled(False)
        self.compare_btn.setStyleSheet(self.primary_button_style())
        self.compare_btn.clicked.connect(self.run_comparison)

        self.cancel_btn = QPushButton("✖ Cancel")
        self.cancel_btn.setFixedHeight(44)
        self.cancel_btn.setVisible(False)
        self.cancel_btn.setStyleSheet(self.secondary_button_style())
        self.cancel_btn.clicked.connect(self.cancel_comparison)

        btn_layout.addWidget(self.compare_btn)
        btn_layout.addWidget(self.cancel_btn)

        layout.addLayout(btn_layout)
        return container
    def run_comparison(self):
        if self.worker and self.worker.isRunning():
            return

        try:
            config = self.build_config()

            self.worker = ComparisonWorker(
                self.df_a,
                self.df_b,
                config,
                self.file_a_path,
                self.file_b_path
            )

            self.worker.progress.connect(self.on_progress)
            self.worker.finished.connect(self.on_finished)
            self.worker.error.connect(self.on_error)

            self.start_time = time.time()

            self.progress_bar.setVisible(True)
            self.progress_label.setText("Starting comparison…")
            self.compare_btn.setEnabled(False)
            self.cancel_btn.setVisible(True)

            self.worker.start()

        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

    def cancel_comparison(self):
        if self.worker and self.worker.isRunning():
            self.worker.terminate()
            self.worker.wait()

        self.reset_ui()
        self.statusBar().showMessage("Comparison cancelled")

    def on_progress(self, message):
        self.progress_label.setText(message)

    def on_finished(self, payload):
        elapsed = time.time() - self.start_time
        self.reset_ui()

        output_path = payload["output_path"]

        QMessageBox.information(
            self,
            "Comparison Complete",
            f"✅ Comparison finished in {elapsed:.2f} seconds.\n\n"
            f"Report generated:\n{output_path}"
        )

        if platform.system() == "Darwin":
            os.system(f"open '{output_path}'")
        elif platform.system() == "Windows":
            os.startfile(output_path)

    def on_error(self, message):
        self.reset_ui()
        QMessageBox.critical(self, "Comparison Failed", message)

    def reset_ui(self):
        self.progress_bar.setVisible(False)
        self.progress_label.setText("")
        self.compare_btn.setEnabled(True)
        self.cancel_btn.setVisible(False)
        self.worker = None

    def primary_button_style(self):
        return f"""
            QPushButton {{
                background: {self.COLOR_PRIMARY};
                color: white;
                font-size: 12pt;
                font-weight: 600;
                padding: 10px 24px;
                border-radius: 8px;
            }}
            QPushButton:hover {{
                background: {self.COLOR_PRIMARY_DARK};
            }}
            QPushButton:disabled {{
                background: {self.COLOR_BORDER};
            }}
        """

    def secondary_button_style(self):
        return f"""
            QPushButton {{
                background: white;
                color: {self.COLOR_TEXT_PRIMARY};
                font-size: 11pt;
                padding: 8px 18px;
                border-radius: 8px;
                border: 2px solid {self.COLOR_BORDER};
            }}
            QPushButton:hover {{
                border-color: {self.COLOR_PRIMARY};
                color: {self.COLOR_PRIMARY};
            }}
        """

    def tertiary_button_style(self):
        return f"""
            QPushButton {{
                background: transparent;
                color: {self.COLOR_PRIMARY};
                font-size: 10pt;
                padding: 6px 12px;
                border: none;
            }}
            QPushButton:hover {{
                text-decoration: underline;
            }}
        """

    def update_compare_button_state(self):
        """Enable Compare button only when both files are loaded"""
        ready = self.df_a is not None and self.df_b is not None
        self.compare_btn.setEnabled(ready)
        self.config_group.setEnabled(ready)

    def select_file(self, which):
        path, _ = QFileDialog.getOpenFileName(
            self,
            f"Select Excel File {which}",
            self.last_directory,
            "Excel Files (*.xlsx *.xls *.xlsm)"
        )
        if path:
            if which == "A":
                self.file_a_display.setText(path)
            else:
                self.file_b_display.setText(path)

    def load_file_path(self, path, which):
        try:
            self.last_directory = str(Path(path).parent)
            self.settings.setValue("last_directory", self.last_directory)

            xls = pd.ExcelFile(path)
            sheets = xls.sheet_names

            if len(sheets) > 1:
                sheet, ok = QInputDialog.getItem(
                    self,
                    "Select Sheet",
                    f"Choose sheet from File {which}:",
                    sheets,
                    0,
                    False
                )
                if not ok:
                    return
            else:
                sheet = sheets[0]

            df = pd.read_excel(path, sheet_name=sheet)

            if which == "A":
                self.file_a_path = path
                self.file_a_sheet = sheet
                self.df_a = df
            else:
                self.file_b_path = path
                self.file_b_sheet = sheet
                self.df_b = df

            self.populate_columns()
            self.update_compare_button_state()

        except Exception as e:
            QMessageBox.critical(self, "File Load Error", str(e))
            self.clear_file(which)

    def clear_file(self, which):
        if which == "A":
            self.file_a_path = None
            self.file_a_sheet = None
            self.df_a = None
            self.file_a_display.clear()
        else:
            self.file_b_path = None
            self.file_b_sheet = None
            self.df_b = None
            self.file_b_display.clear()

        self.populate_columns()

    def populate_columns(self):
        # Reset UI
        for cb in self.key_checkboxes:
            cb.deleteLater()
        self.key_checkboxes.clear()

        if self.df_a is None or self.df_b is None:
            self.key_placeholder.setVisible(True)
            self.key_scroll.setVisible(False)
            self.key_filter.setVisible(False)
            self.select_all_btn.setVisible(False)
            self.deselect_all_btn.setVisible(False)
            self.key_count_label.setVisible(False)
            self.tiebreaker_combo.clear()
            return

        common_cols = sorted(set(self.df_a.columns) & set(self.df_b.columns))

        if not common_cols:
            QMessageBox.warning(
                self,
                "No Common Columns",
                "The two files have no matching column names."
            )
            return

        self.key_placeholder.setVisible(False)
        self.key_scroll.setVisible(True)
        self.key_filter.setVisible(True)
        self.select_all_btn.setVisible(True)
        self.deselect_all_btn.setVisible(True)
        self.key_count_label.setVisible(True)

        for i, col in enumerate(common_cols):
            cb = QCheckBox(col)
            cb.setStyleSheet(self.modern_checkbox_style())
            cb.stateChanged.connect(self.update_key_count)

            row, col_pos = divmod(i, 2)
            self.key_grid.addWidget(cb, row, col_pos)
            self.key_checkboxes.append(cb)

            self.tiebreaker_combo.addItem(col, col)

        self.update_key_count()

    def update_key_count(self):
        selected = sum(cb.isChecked() for cb in self.key_checkboxes)
        self.key_count_label.setText(f"Selected keys: {selected}")

    def toggle_all_keys(self, checked):
        for cb in self.key_checkboxes:
            cb.setChecked(checked)

    def filter_key_columns(self, text):
        text = text.lower()
        for cb in self.key_checkboxes:
            cb.setVisible(text in cb.text().lower())
    def build_config(self):
        if self.mode_key_based.isChecked():
            keys = [cb.text() for cb in self.key_checkboxes if cb.isChecked()]
            if not keys:
                raise ValueError("Please select at least one key column.")

            return ComparisonConfig(
                alignment_method=AlignmentMethod.KEY,
                key_columns=keys,
                tiebreaker_column=self.tiebreaker_combo.currentData(),
                case_sensitive=self.case_sensitive.isChecked(),
                trim_whitespace=self.trim_whitespace.isChecked()
            )

        return ComparisonConfig(
            alignment_method=AlignmentMethod.POSITION,
            case_sensitive=self.case_sensitive.isChecked(),
            trim_whitespace=self.trim_whitespace.isChecked()
        )

    def load_settings(self):
        pass

def main():
    app = QApplication(sys.argv)
    app.setApplicationName("GridKit")
    app.setOrganizationName("ExcelCompTool")

    window = ExcelComparisonGUI()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()

