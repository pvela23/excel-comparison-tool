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
    QComboBox, QInputDialog, QFrame, QSizePolicy, QRadioButton, QButtonGroup
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
        # App Icon
        if os.path.exists("GridKit.ico"):
            self.setWindowIcon(QIcon("GridKit.ico"))

        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setSpacing(12)
        main_layout.setContentsMargins(16, 16, 16, 16)

        # Modern header
        title = QLabel("GridKit")
        title.setFont(self.ui_font(size=24, bold=True))
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet(f"""
            color: {self.COLOR_TEXT_PRIMARY};
            padding: 4px;
        """)
        main_layout.addWidget(title)

        subtitle = QLabel("Compare two Excel files and highlight differences")
        subtitle.setAlignment(Qt.AlignCenter)
        subtitle.setStyleSheet(f"""
            color: {self.COLOR_TEXT_SECONDARY};
            font-size: 11pt;
            padding-bottom: 4px;
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
        
        # Compare section (Buttons + Progress) moved inside content layout
        # main_layout.addWidget(self.create_compare_section()) <- Removed from bottom
        
        # Add compare section directly below config
        content_layout.addWidget(self.create_compare_section())
        content_layout.addStretch()

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
                padding: 0 0;
                left: 0;
                top: 0;
                background: transparent;
                border: none;
            }}
        """

    def flat_card_style(self):
        return f"""
            QFrame {{
                background: white;
                border: 1px solid {self.COLOR_BORDER};
                border-radius: 8px;
            }}
        """

    def mode_card_style(self):
        return f"""
            QFrame {{
                background: {self.COLOR_BG_LIGHT};
                border: 1px solid {self.COLOR_BORDER};
                border-radius: 8px;
                padding: 12px;
            }}
        """

    # ---------- File Section ----------
    def create_file_section(self):
        group = QGroupBox("📁 1. Select Files")
        group.setStyleSheet(self.card_style())
        layout = QVBoxLayout(group)
        layout.setSpacing(6)
        layout.setContentsMargins(12, 16, 12, 12)

        # Single grid for both files to ensure perfect alignment
        # Columns: [Label] [Input] [Button]
        grid_layout = QGridLayout()
        grid_layout.setSpacing(6)
        
        # --- File A ---
        lbl_a = QLabel("File A:")
        lbl_a.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        lbl_a.setStyleSheet(f"font-size: 11pt; font-weight: 600; color: {self.COLOR_TEXT_PRIMARY};")
       
        self.file_a_display = QLineEdit()
        self.file_a_display.setPlaceholderText("Drag & drop, browse, or paste file path...")
        self.file_a_display.setStyleSheet(f"""
            QLineEdit {{
                padding: 6px 8px;
                font-size: 11pt;
                background: #f8f9fa;
                color: {self.COLOR_TEXT_PRIMARY};
                border: 2px solid {self.COLOR_BORDER};
                border-radius: 6px;
            }}
            QLineEdit:focus {{
                border-color: {self.COLOR_PRIMARY};
                background: white;
            }}
        """)
        self.file_a_display.textChanged.connect(lambda: self.on_file_path_changed("A"))
       
        btn_a = QPushButton("Browse")
        btn_a.setFixedWidth(80)
        btn_a.setStyleSheet(self.secondary_button_style())
        btn_a.clicked.connect(lambda: self.select_file("A"))

        grid_layout.addWidget(lbl_a, 0, 0)
        grid_layout.addWidget(self.file_a_display, 0, 1)
        grid_layout.addWidget(btn_a, 0, 2)
        
        # Tip A
        tip_a = QLabel("Original (before) file")
        tip_a.setStyleSheet(f"font-size: 10pt; color: {self.COLOR_TEXT_SECONDARY}; padding-left: 4px; font-style: italic;")
        grid_layout.addWidget(tip_a, 1, 1)

        # --- File B ---
        lbl_b = QLabel("File B:")
        lbl_b.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        lbl_b.setStyleSheet(f"font-size: 11pt; font-weight: 600; color: {self.COLOR_TEXT_PRIMARY};")
       
        self.file_b_display = QLineEdit()
        self.file_b_display.setPlaceholderText("Drag & drop, browse, or paste file path...")
        self.file_b_display.setStyleSheet(f"""
            QLineEdit {{
                padding: 6px 8px;
                font-size: 11pt;
                background: #f0f8ff;
                color: {self.COLOR_TEXT_PRIMARY};
                border: 2px solid {self.COLOR_BORDER};
                border-radius: 6px;
            }}
            QLineEdit:focus {{
                border-color: {self.COLOR_PRIMARY};
                background: white;
            }}
        """)
        self.file_b_display.textChanged.connect(lambda: self.on_file_path_changed("B"))
       
        btn_b = QPushButton("Browse")
        btn_b.setFixedWidth(80)
        btn_b.setStyleSheet(self.secondary_button_style())
        btn_b.clicked.connect(lambda: self.select_file("B"))

        grid_layout.addWidget(lbl_b, 2, 0)
        grid_layout.addWidget(self.file_b_display, 2, 1)
        grid_layout.addWidget(btn_b, 2, 2)
        
        # Tip B
        tip_b = QLabel("Updated (after) file")
        tip_b.setStyleSheet(f"font-size: 10pt; color: {self.COLOR_TEXT_SECONDARY}; padding-left: 4px; font-style: italic;")
        grid_layout.addWidget(tip_b, 3, 1) # Skip a row for spacing if needed or put right below

        # Set column stretch
        grid_layout.setColumnStretch(1, 1)
        grid_layout.setColumnStretch(0, 0)
        grid_layout.setColumnStretch(2, 0)
        
        layout.addLayout(grid_layout)
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
        # Header for the section (Label instead of GroupBox title for cleaner look)
        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setSpacing(16)
        layout.setContentsMargins(12, 16, 12, 12)
        
        header = QLabel("🔧 2. Configure Comparison")
        header.setStyleSheet(f"font-size: 13pt; font-weight: bold; color: {self.COLOR_TEXT_PRIMARY};")
        layout.addWidget(header)
        
        # === Row Matching Mode ===
        mode_label = QLabel("How are rows identified?")
        mode_label.setStyleSheet(f"font-size: 11pt; font-weight: 600; color: {self.COLOR_TEXT_PRIMARY};")
        layout.addWidget(mode_label)

        # Mode Selection: Use Toggle-like or just Radio buttons horizontally
        # User requested look: [Switch] Key-Based
        # We'll use horizontal layout for modes since it's cleaner now
        
        self.mode_group = QButtonGroup(self)
        
        mode_layout = QHBoxLayout()
        mode_layout.setSpacing(24)
        
        self.mode_key_based = QRadioButton("Key-Based")
        self.mode_key_based.setChecked(True)
        self.mode_key_based.setEnabled(False) 
        self.mode_key_based.setStyleSheet(self.modern_radio_style())
        self.mode_key_based.toggled.connect(self.on_mode_changed)
        self.mode_group.addButton(self.mode_key_based)
        mode_layout.addWidget(self.mode_key_based)
       
        self.mode_position_based = QRadioButton("Position-Based")
        self.mode_position_based.setEnabled(False)
        self.mode_position_based.setStyleSheet(self.modern_radio_style())
        self.mode_position_based.toggled.connect(self.on_mode_changed)
        self.mode_group.addButton(self.mode_position_based)
        mode_layout.addWidget(self.mode_position_based)
        
        mode_layout.addStretch()
        layout.addLayout(mode_layout)

        # Position Info
        self.position_info = QLabel("ℹ️ Compares files row-by-row (Row 1 to Row 1).")
        self.position_info.setStyleSheet(f"font-size: 10pt; color: {self.COLOR_TEXT_SECONDARY}; margin-left: 24px;")
        self.position_info.setVisible(False)
        layout.addWidget(self.position_info)

        # === Key Columns ===
        self.key_frame = QWidget()
        self.key_frame.setVisible(True) # Visible by default with placeholder
        key_layout = QVBoxLayout(self.key_frame)
        key_layout.setSpacing(8)
        key_layout.setContentsMargins(0, 8, 0, 0)
        
        lbl_keys = QLabel("Key Columns")
        lbl_keys.setStyleSheet(f"font-size: 11pt; font-weight: 600; color: {self.COLOR_TEXT_PRIMARY};")
        key_layout.addWidget(lbl_keys)
        
        # Dynamic Scroll Area for Keys
        self.key_scroll = QScrollArea()
        self.key_scroll.setWidgetResizable(True)
        self.key_scroll.setFrameShape(QFrame.Shape.NoFrame)
        # Style to look like a simple list area
        self.key_scroll.setStyleSheet(f"""
            QScrollArea {{
                background: transparent;
                border: 1px dashed {self.COLOR_BORDER};
                border-radius: 6px;
            }}
            QWidget {{ background: transparent; }}
        """)
        
        self.key_container = QWidget()
        self.key_grid = QVBoxLayout(self.key_container) 
        self.key_grid.setSpacing(4)
        self.key_grid.setContentsMargins(8, 8, 8, 8)
        self.key_grid.setAlignment(Qt.AlignmentFlag.AlignTop)
        
        self.key_scroll.setWidget(self.key_container)
        
        # Placeholder Label inside the scroll area initially
        self.key_placeholder_label = QLabel("Load files to select keys")
        self.key_placeholder_label.setAlignment(Qt.AlignCenter)
        self.key_placeholder_label.setStyleSheet(f"color: {self.COLOR_TEXT_TERTIARY}; padding: 20px;")
        
        self.key_grid.addWidget(self.key_placeholder_label)
        
        # Initial height small
        self.key_scroll.setFixedHeight(80) 
        key_layout.addWidget(self.key_scroll)
        
        layout.addWidget(self.key_frame)

        # === Advanced Options ===
        # Simple toggle button like user asked
        self.advanced_toggle = QPushButton("▼ Advanced Options")
        self.advanced_toggle.setVisible(False)
        self.advanced_toggle.setCursor(Qt.CursorShape.PointingHandCursor)
        self.advanced_toggle.setStyleSheet(f"""
            QPushButton {{
                text-align: left;
                padding: 8px 0;
                font-size: 11pt;
                color: {self.COLOR_PRIMARY};
                background: transparent;
                border: none;
                font-weight: 600;
            }}
            QPushButton:hover {{ text-decoration: underline; }}
        """)
        self.advanced_toggle.clicked.connect(self.toggle_advanced_options)
        layout.addWidget(self.advanced_toggle)

        # Advanced Container
        self.advanced_container = QWidget()
        adv_layout = QGridLayout(self.advanced_container)
        adv_layout.setSpacing(12)
        adv_layout.setContentsMargins(0, 0, 0, 0)
        
        # Tiebreaker
        adv_layout.addWidget(QLabel("Tiebreaker:"), 0, 0)
        self.tiebreaker_combo = QComboBox()
        self.tiebreaker_combo.setStyleSheet(f"border: 1px solid {self.COLOR_BORDER}; padding: 4px; border-radius: 4px;")
        adv_layout.addWidget(self.tiebreaker_combo, 0, 1)
        
        # Checkboxes
        self.case_sensitive = QCheckBox("Case Sensitive")
        self.trim_whitespace = QCheckBox("Trim Whitespace")
        self.trim_whitespace.setChecked(True)
        adv_layout.addWidget(self.case_sensitive, 1, 0)
        adv_layout.addWidget(self.trim_whitespace, 1, 1)
        
        self.advanced_container.setVisible(False)
        layout.addWidget(self.advanced_container)
        
        # Tiebreaker tip
        self.tiebreaker_tip = QLabel("💡 Use sort col for order mismatch")
        self.tiebreaker_tip.setVisible(False)
        self.tiebreaker_tip.setStyleSheet(f"color: {self.COLOR_TEXT_SECONDARY}; font-size: 9pt;")
        layout.addWidget(self.tiebreaker_tip)

        # Store for access
        self.config_group = container 
        return self.config_group

    def modern_radio_style(self):
        return f"""
            QRadioButton {{
                font-size: 12pt;
                color: {self.COLOR_TEXT_PRIMARY};
                spacing: 8px;
            }}
            QRadioButton::indicator {{
                width: 16px;
                height: 16px;
                border-radius: 8px;
                border: 2px solid {self.COLOR_BORDER};
                background: white;
            }}
            QRadioButton::indicator:hover {{
                border-color: {self.COLOR_PRIMARY};
            }}
            QRadioButton::indicator:checked {{
                background-color: {self.COLOR_PRIMARY};
                border-color: {self.COLOR_PRIMARY};
            }}
        """

    def modern_checkbox_style(self):
        return f"""
            QCheckBox {{
                font-size: 12pt;
                color: {self.COLOR_TEXT_PRIMARY};
                spacing: 8px;
            }}

            QCheckBox::indicator {{
                width: 16px;
                height: 16px;
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

    def primary_button_style(self):
        return f"""
            QPushButton {{
                background: {self.COLOR_PRIMARY};
                color: white;
                font-size: 12pt;
                font-weight: 600;
                padding: 8px 20px;
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
                padding: 6px 12px;
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
                font-size: 11pt;
                padding: 4px 8px;
                border: none;
            }}
            QPushButton:hover {{
                text-decoration: underline;
            }}
        """



    def toggle_advanced_options(self):
        """Toggle advanced options visibility"""
        self.advanced_expanded = not self.advanced_expanded
        self.advanced_container.setVisible(self.advanced_expanded)
        self.advanced_toggle.setText("▲ Advanced Options" if self.advanced_expanded else "▼ Advanced Options")
    
    def on_mode_changed(self):
        """Handle mode change with radio button behavior"""
        if self.mode_key_based.isChecked():
            self.key_frame.setVisible(True)
            self.position_info.setVisible(False)
        else:
            self.key_frame.setVisible(False)
            self.position_info.setVisible(True)

    def on_tiebreaker_changed(self):
        """Handle tiebreaker selection"""
        tiebreaker = self.tiebreaker_combo.currentData()
        self.tiebreaker_tip.setVisible(tiebreaker is not None)



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
        layout.setSpacing(8)

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
        btn_layout.addStretch() # Right Aligned

        self.compare_btn = QPushButton("Compare →")
        self.compare_btn.setFixedHeight(48)
        self.compare_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.compare_btn.setEnabled(False)
        self.compare_btn.setStyleSheet(f"""
            QPushButton {{
                background: {self.COLOR_PRIMARY};
                color: white;
                font-size: 13pt;
                font-weight: 600;
                border-radius: 8px;
                padding: 0 24px;
            }}
            QPushButton:hover {{ background: {self.COLOR_PRIMARY_DARK}; }}
            QPushButton:disabled {{ background: {self.COLOR_BORDER}; }}
        """)
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
        
        # Enable specific config elements
        self.mode_key_based.setEnabled(ready)
        self.mode_position_based.setEnabled(ready)
        
        # Show Advanced Toggle based on readiness
        self.advanced_toggle.setVisible(ready)
        
        # Key Frame is always visible in Key mode now (with placeholder if not ready)
        if self.mode_key_based.isChecked():
            self.key_frame.setVisible(True)
        else:
            self.key_frame.setVisible(False)

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
            # Restore placeholder
            while self.key_grid.count():
                item = self.key_grid.takeAt(0)
                if item.widget(): item.widget().deleteLater()
            
            self.key_placeholder_label = QLabel("Load files to select keys")
            self.key_placeholder_label.setAlignment(Qt.AlignCenter)
            self.key_placeholder_label.setStyleSheet(f"color: {self.COLOR_TEXT_TERTIARY}; padding: 20px;")
            self.key_grid.addWidget(self.key_placeholder_label)
            self.key_scroll.setFixedHeight(80)
            
            self.key_scroll.setVisible(True)
            self.key_filter.setVisible(False)
            self.select_all_btn.setVisible(False)
            self.deselect_all_btn.setVisible(False)
            self.key_count_label.setVisible(False)
            self.tiebreaker_combo.clear()
            return

        # Preserve order from File A
        cols_a = list(self.df_a.columns)
        cols_b_set = set(self.df_b.columns)
        common_cols = [c for c in cols_a if c in cols_b_set]

        if not common_cols:
            QMessageBox.warning(
                self,
                "No Common Columns",
                "The two files have no matching column names."
            )
            return

        # Calculate dynamic height
        # Row height roughly 28px. 4 rows = 112px. + Padding ~20px = 132px.
        num_keys = len(common_cols)
        
        # We need to clear grid properly (widgets are not automatically deleted with clear)
        # But we called clear() before which is just for the list self.key_checkboxes
        # Logic to clear layout items:
        while self.key_grid.count():
            item = self.key_grid.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()
        
        self.key_scroll.setVisible(True)
        self.key_placeholder_label.setVisible(False) # Hide placeholder
        
        # Calculate size: 
        row_height = 32 # approx
        spacing = 4
        # If <= 4, fit content. If > 4, limit.
        if num_keys <= 4:
            total_h = (num_keys * row_height) + (num_keys * spacing) + 20
            self.key_scroll.setFixedHeight(total_h)
            self.key_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        else:
            total_h = (4 * row_height) + (4 * spacing) + 20
            self.key_scroll.setFixedHeight(total_h)
            self.key_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)

        self.key_checkboxes = []
        # Populate
        self.tiebreaker_combo.addItem("(None)", None)
        for i, col in enumerate(common_cols):
            cb = QCheckBox(col)
            cb.setStyleSheet("font-size: 11pt; padding: 4px;")
            cb.stateChanged.connect(self.update_key_count)
            self.key_grid.addWidget(cb)
            self.key_checkboxes.append(cb)
            self.tiebreaker_combo.addItem(col, col)
            
        self.key_grid.addStretch() # Push items to top

        if self.tiebreaker_combo.count() > 0:
            self.tiebreaker_combo.setCurrentIndex(0)

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

            method = AlignmentMethod.POSITION
            if self.tiebreaker_combo.currentData():
                method = AlignmentMethod.SECONDARY_SORT

            return ComparisonConfig(
                alignment_method=method,
                key_columns=keys,
                secondary_sort_column=self.tiebreaker_combo.currentData(),
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
