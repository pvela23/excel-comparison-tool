"""
Enhanced GUI for Excel Comparison Tool
Built with PySide6 - Compact & Feature-Rich
"""

import sys
from pathlib import Path
from datetime import datetime
import time
import os
import platform
import pandas as pd

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QGroupBox, QCheckBox,
    QProgressBar, QMessageBox, QScrollArea, QGridLayout, QLineEdit,
    QComboBox, QInputDialog
)
from PySide6.QtCore import Qt, QThread, Signal, QSettings
from PySide6.QtGui import QFont, QAction, QKeySequence, QDragEnterEvent, QDropEvent

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
# Main GUI
# =========================
class ExcelComparisonGUI(QMainWindow):

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
        self.setWindowTitle("Excel Comparison Tool v1.0")
        self.setMinimumSize(900, 600)
        self.resize(1000, 700)

        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setSpacing(10)
        main_layout.setContentsMargins(15, 15, 15, 15)

        # Title
        title = QLabel("Excel Comparison Tool")
        title.setFont(self.ui_font(size=16, bold=True))
        title.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title)

        subtitle = QLabel("Compare two Excel files using key-based or position-based matching")
        subtitle.setAlignment(Qt.AlignCenter)
        subtitle.setStyleSheet("color: gray; font-size: 9pt;")
        main_layout.addWidget(subtitle)

        # Sections
        main_layout.addWidget(self.create_file_section())
        main_layout.addWidget(self.create_config_section())
        main_layout.addWidget(self.create_compare_section())
        main_layout.addStretch()

        self.statusBar().showMessage("Ready - Drag & drop Excel files or use Browse buttons")
       
        # Keyboard shortcuts
        self.setup_shortcuts()

    def ui_font(self, size=9, bold=False):
        font = QFont()
        font.setPointSize(size)
        if bold:
            font.setWeight(QFont.Weight.Bold)
        return font

    def setup_shortcuts(self):
        """Setup keyboard shortcuts"""
        compare_action = QAction("Compare", self)
        compare_action.setShortcut(QKeySequence("Ctrl+Return"))
        compare_action.triggered.connect(self.run_comparison)
        self.addAction(compare_action)
       
        # Tooltip for compare button
        self.compare_btn.setToolTip("Click or press Ctrl+Enter to compare")

    # ---------- File Section ----------
    def create_file_section(self):
        group = QGroupBox("📁 1. Select Files")
        group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                font-size: 10pt;
                padding-top: 12px;
                margin-top: 8px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 8px;
                padding: 0 4px;
            }
        """)
        layout = QGridLayout(group)
        layout.setSpacing(8)
        layout.setContentsMargins(10, 15, 10, 10)

        # File A
        lbl_a = QLabel("File A:")
        lbl_a.setStyleSheet("font-weight: normal; font-size: 9pt;")
       
        self.file_a_display = QLineEdit()
        self.file_a_display.setPlaceholderText("No file selected (drag & drop, browse, or paste path)")
        self.file_a_display.setStyleSheet("""
            QLineEdit {
                padding: 6px;
                font-size: 9pt;
                background-color: white;
                border: 1px solid #CCC;
                border-radius: 3px;
            }
        """)
        self.file_a_display.textChanged.connect(lambda: self.on_file_path_changed("A"))
       
        btn_a = QPushButton("Browse...")
        btn_a.setFixedWidth(90)
        btn_a.setFixedHeight(28)
        btn_a.setStyleSheet(self.button_style())
        btn_a.clicked.connect(lambda: self.select_file("A"))

        # File B
        lbl_b = QLabel("File B:")
        lbl_b.setStyleSheet("font-weight: normal; font-size: 9pt;")
       
        self.file_b_display = QLineEdit()
        self.file_b_display.setPlaceholderText("No file selected (drag & drop, browse, or paste path)")
        self.file_b_display.setStyleSheet("""
            QLineEdit {
                padding: 6px;
                font-size: 9pt;
                background-color: white;
                border: 1px solid #CCC;
                border-radius: 3px;
            }
        """)
        self.file_b_display.textChanged.connect(lambda: self.on_file_path_changed("B"))
       
        btn_b = QPushButton("Browse...")
        btn_b.setFixedWidth(90)
        btn_b.setFixedHeight(28)
        btn_b.setStyleSheet(self.button_style())
        btn_b.clicked.connect(lambda: self.select_file("B"))

        layout.addWidget(lbl_a, 0, 0)
        layout.addWidget(self.file_a_display, 0, 1)
        layout.addWidget(btn_a, 0, 2)

        layout.addWidget(lbl_b, 1, 0)
        layout.addWidget(self.file_b_display, 1, 1)
        layout.addWidget(btn_b, 1, 2)

        layout.setColumnStretch(1, 1)
        return group

    def on_file_path_changed(self, which):
        """Handle manual file path entry"""
        if which == "A":
            path = self.file_a_display.text().strip()
        else:
            path = self.file_b_display.text().strip()
       
        if path and Path(path).exists() and Path(path).is_file():
            self.load_file_path(path, which)

    # ---------- Config Section ----------
    def create_config_section(self):
        self.config_group = QGroupBox("⚙️ 2. Configure Comparison")
        self.config_group.setEnabled(False)
        self.config_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                font-size: 10pt;
                padding-top: 12px;
                margin-top: 8px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 8px;
                padding: 0 4px;
            }
        """)
        layout = QVBoxLayout(self.config_group)
        layout.setSpacing(10)
        layout.setContentsMargins(10, 15, 10, 10)

        # ---- Comparison Mode ----
        mode_layout = QHBoxLayout()
        mode_layout.setSpacing(8)
       
        mode_label = QLabel("Comparison Mode:")
        mode_label.setStyleSheet("font-weight: normal; font-size: 9pt;")
        mode_layout.addWidget(mode_label)
       
        self.mode_key_based = QCheckBox("Key-Based (Row Matching)")
        self.mode_key_based.setChecked(True)
        self.mode_key_based.setStyleSheet("font-size: 9pt; font-weight: bold;")
        self.mode_key_based.toggled.connect(self.on_mode_changed)
       
        self.mode_position_based = QCheckBox("Position-Based (Row 1 → Row 1)")
        self.mode_position_based.setStyleSheet("font-size: 9pt; font-weight: bold;")
        self.mode_position_based.toggled.connect(self.on_mode_changed)
       
        mode_layout.addWidget(self.mode_key_based)
        mode_layout.addWidget(self.mode_position_based)
        mode_layout.addStretch()
       
        layout.addLayout(mode_layout)

        # ---- Key Columns Section ----
        self.key_section = QWidget()
        key_section_layout = QVBoxLayout(self.key_section)
        key_section_layout.setSpacing(8)
        key_section_layout.setContentsMargins(0, 0, 0, 0)
       
        key_header = QLabel("Select Key Columns (unique row identifier):")
        key_header.setStyleSheet("font-weight: normal; font-size: 9pt;")
        key_section_layout.addWidget(key_header)

        # Select All / Deselect All buttons
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(6)
       
        self.select_all_btn = QPushButton("Select All")
        self.select_all_btn.setFixedHeight(26)
        self.select_all_btn.setStyleSheet(self.small_button_style())
        self.select_all_btn.clicked.connect(lambda: self.toggle_all_keys(True))
       
        self.deselect_all_btn = QPushButton("Deselect All")
        self.deselect_all_btn.setFixedHeight(26)
        self.deselect_all_btn.setStyleSheet(self.small_button_style())
        self.deselect_all_btn.clicked.connect(lambda: self.toggle_all_keys(False))
       
        btn_layout.addWidget(self.select_all_btn)
        btn_layout.addWidget(self.deselect_all_btn)
        btn_layout.addStretch()
        key_section_layout.addLayout(btn_layout)

        # Filter
        self.key_filter = QLineEdit()
        self.key_filter.setPlaceholderText("🔍 Filter columns...")
        self.key_filter.setFixedHeight(28)
        self.key_filter.setStyleSheet("""
            QLineEdit {
                padding: 5px;
                font-size: 9pt;
                border: 1px solid #CCC;
                border-radius: 3px;
            }
        """)
        self.key_filter.textChanged.connect(self.filter_key_columns)
        key_section_layout.addWidget(self.key_filter)

        # Scroll area
        self.key_scroll = QScrollArea()
        self.key_scroll.setWidgetResizable(True)
        self.key_scroll.setMinimumHeight(120)
        self.key_scroll.setMaximumHeight(180)
        self.key_scroll.setStyleSheet("""
            QScrollArea {
                border: 1px solid #CCC;
                border-radius: 3px;
                background-color: white;
            }
            QCheckBox {
                font-size: 9pt;
                padding: 2px;
            }
        """)

        self.key_container = QWidget()
        self.key_grid = QGridLayout(self.key_container)
        self.key_grid.setSpacing(6)
        self.key_grid.setContentsMargins(8, 8, 8, 8)

        self.key_scroll.setWidget(self.key_container)
        key_section_layout.addWidget(self.key_scroll)

        # Key count label
        self.key_count_label = QLabel("")
        self.key_count_label.setStyleSheet("font-size: 8pt; color: gray; padding: 2px;")
        key_section_layout.addWidget(self.key_count_label)
       
        layout.addWidget(self.key_section)

        # ---- Position-Based Info ----
        self.position_info = QLabel(
            "ℹ️ Position-based mode compares files row-by-row without keys.\n"
            "Row 1 in File A is compared to Row 1 in File B, etc."
        )
        self.position_info.setStyleSheet("""
            QLabel {
                font-size: 8pt;
                color: #555;
                background-color: #F0F8FF;
                border: 1px solid #B0D4FF;
                border-radius: 3px;
                padding: 8px;
            }
        """)
        self.position_info.setWordWrap(True)
        self.position_info.setVisible(False)
        layout.addWidget(self.position_info)

        # ---- Options ----
        options_layout = QGridLayout()
        options_layout.setSpacing(8)
        options_layout.setColumnStretch(1, 1)
       
        # Tiebreaker column selector (only for Key-Based mode)
        self.tiebreaker_label = QLabel("Tiebreaker Column:")
        self.tiebreaker_label.setStyleSheet("font-weight: normal; font-size: 9pt;")
       
        self.tiebreaker_combo = QComboBox()
        self.tiebreaker_combo.setFixedHeight(28)
        self.tiebreaker_combo.setStyleSheet("""
            QComboBox {
                padding: 5px;
                font-size: 9pt;
                border: 1px solid #CCC;
                border-radius: 3px;
            }
        """)
        
        options_layout.addWidget(self.tiebreaker_label, 0, 0, Qt.AlignmentFlag.AlignRight)
        options_layout.addWidget(self.tiebreaker_combo, 0, 1)
       
        self.case_sensitive = QCheckBox("Case Sensitive")
        self.case_sensitive.setStyleSheet("font-size: 9pt;")
       
        self.trim_whitespace = QCheckBox("Trim Whitespace")
        self.trim_whitespace.setChecked(True)
        self.trim_whitespace.setStyleSheet("font-size: 9pt;")

        options_layout.addWidget(self.case_sensitive, 2, 1)
        options_layout.addWidget(self.trim_whitespace, 3, 1)
       
        layout.addLayout(options_layout)

        return self.config_group
   

    
    def on_mode_changed(self):
        """Handle comparison mode change - now uses radio button logic"""
        sender = self.sender()
        
        if sender == self.mode_key_based and self.mode_key_based.isChecked():
            # Key-based mode selected
            self.mode_position_based.blockSignals(True)
            self.mode_position_based.setChecked(False)
            self.mode_position_based.blockSignals(False)
            self.key_section.setVisible(True)
            self.position_info.setVisible(False)
            self.tiebreaker_label.setVisible(True)
            self.tiebreaker_combo.setVisible(True)
            
        elif sender == self.mode_position_based and self.mode_position_based.isChecked():
            # Position-based mode selected
            self.mode_key_based.blockSignals(True)
            self.mode_key_based.setChecked(False)
            self.mode_key_based.blockSignals(False)
            self.key_section.setVisible(False)
            self.position_info.setVisible(True)
            self.tiebreaker_label.setVisible(False)
            self.tiebreaker_combo.setVisible(False)
            
        elif not self.mode_key_based.isChecked() and not self.mode_position_based.isChecked():
            # If user unchecks one, re-check it (radio button behavior)
            if sender == self.mode_key_based:
                self.mode_key_based.blockSignals(True)
                self.mode_key_based.setChecked(True)
                self.mode_key_based.blockSignals(False)
            else:
                self.mode_position_based.blockSignals(True)
                self.mode_position_based.setChecked(True)
                self.mode_position_based.blockSignals(False)

    # ---------- Compare Section ----------
    def create_compare_section(self):
        group = QGroupBox("🔍 3. Start Comparison")
        group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                font-size: 10pt;
                padding-top: 12px;
                margin-top: 8px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 8px;
                padding: 0 4px;
            }
        """)
        layout = QVBoxLayout(group)
        layout.setSpacing(8)
        layout.setContentsMargins(10, 15, 10, 10)

        self.compare_btn = QPushButton("🔍 Compare Files")
        self.compare_btn.setMinimumHeight(42)
        self.compare_btn.setEnabled(False)
        self.compare_btn.setStyleSheet("""
            QPushButton {
                background-color: #5B7FB8;
                color: white;
                font-size: 11pt;
                font-weight: bold;
                border-radius: 4px;
                border: none;
            }
            QPushButton:hover {
                background-color: #4A6A9E;
            }
            QPushButton:disabled {
                background-color: #CCC;
                color: #888;
            }
        """)
        self.compare_btn.clicked.connect(self.run_comparison)

        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setFixedHeight(22)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid #CCC;
                border-radius: 3px;
                text-align: center;
                font-size: 8pt;
            }
            QProgressBar::chunk {
                background-color: #5B7FB8;
            }
        """)

        layout.addWidget(self.compare_btn)
        layout.addWidget(self.progress_bar)
        return group

    # ---------- Styles ----------
    def button_style(self):
        return """
            QPushButton {
                padding: 6px 12px;
                font-size: 9pt;
                background-color: #F0F0F0;
                border: 1px solid #CCC;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #E0E0E0;
            }
        """

    def small_button_style(self):
        return """
            QPushButton {
                padding: 4px 10px;
                font-size: 8pt;
                background-color: #F8F8F8;
                border: 1px solid #CCC;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #E8E8E8;
            }
        """

    # ---------- File Handling ----------
    def select_file(self, which):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Excel File",
            self.last_directory,
            "Excel Files (*.xlsx *.xls *.xlsm)"
        )
        if not path:
            return
       
        # Remember directory
        self.last_directory = str(Path(path).parent)
        
        # Set the path in the display field
        if which == "A":
            self.file_a_display.setText(path)
        else:
            self.file_b_display.setText(path)

    def load_file_path(self, path, which):
        """Load a file given its path"""
        try:
            path_obj = Path(path)
            if not path_obj.exists():
                QMessageBox.warning(self, "File Not Found", f"File not found: {path}")
                return
            
            # Get sheet names
            excel_file = pd.ExcelFile(path)
            sheet_names = excel_file.sheet_names
           
            # If multiple sheets, let user choose
            sheet_name = sheet_names[0]  # Default to first sheet
            if len(sheet_names) > 1:
                sheet_name, ok = QInputDialog.getItem(
                    self, "Select Sheet",
                    f"File has {len(sheet_names)} sheets. Select one:",
                    sheet_names, 0, False
                )
                if not ok:
                    return
           
            # Load with string dtype to prevent conversions
            df = pd.read_excel(path, sheet_name=sheet_name, dtype=str)
           
            # Validate
            if df.empty:
                QMessageBox.warning(
                    self, "Empty File",
                    f"The selected sheet appears to be empty.\n\nFile: {path_obj.name}"
                )
                return
           
            # Guardrail on file size
            if len(df) > 500_000:
                reply = QMessageBox.question(
                    self, "Large File Warning",
                    f"This file has {len(df):,} rows, which may consume significant memory.\n\n"
                    "For files over 500,000 rows, comparison may be slow.\n\n"
                    "Continue anyway?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.No
                )
                if reply == QMessageBox.StandardButton.No:
                    return
           
            if which == "A":
                self.file_a_path = path
                self.file_a_sheet = sheet_name
                self.df_a = df
                self.file_a_display.setText(path)
                self.file_a_display.setToolTip(f"File: {path}\nRows: {len(df):,}\nColumns: {len(df.columns)}")
                self.statusBar().showMessage(
                    f"✅ File A loaded: {len(df):,} rows, {len(df.columns)} columns"
                )
            else:
                self.file_b_path = path
                self.file_b_sheet = sheet_name
                self.df_b = df
                self.file_b_display.setText(path)
                self.file_b_display.setToolTip(f"File: {path}\nRows: {len(df):,}\nColumns: {len(df.columns)}")
                self.statusBar().showMessage(
                    f"✅ File B loaded: {len(df):,} rows, {len(df.columns)} columns"
                )

            if self.df_a is not None and self.df_b is not None:
                common_cols = [col for col in self.df_a.columns if col in self.df_b.columns]
               
                if not common_cols:
                    QMessageBox.warning(
                        self, "No Common Columns",
                        "These files have no columns in common!\n\n"
                        f"File A columns: {', '.join(list(self.df_a.columns)[:5])}...\n"
                        f"File B columns: {', '.join(list(self.df_b.columns)[:5])}..."
                    )
                    return
               
                self.update_key_column_options(common_cols)
                self.config_group.setEnabled(True)
                self.compare_btn.setEnabled(True)

        except FileNotFoundError:
            QMessageBox.critical(self, "File Not Found", f"Could not find the file:\n\n{path}")
        except PermissionError:
            QMessageBox.critical(
                self, "Permission Denied",
                f"Cannot access the file (it may be open in Excel):\n\n{path}"
            )
        except ValueError as e:
            QMessageBox.critical(self, "Invalid File Format", f"Invalid Excel file:\n\n{path}")
        except Exception as e:
            QMessageBox.critical(
                self, "Error Loading File",
                f"An unexpected error occurred:\n\n{path}\n\nError: {str(e)}"
            )

    # ---------- Drag & Drop ----------
    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        files = [u.toLocalFile() for u in event.mimeData().urls()]
        excel_files = [f for f in files if f.endswith(('.xlsx', '.xls', '.xlsm'))]
       
        if len(excel_files) >= 2:
            self.file_a_display.setText(excel_files[0])
            self.file_b_display.setText(excel_files[1])
            QMessageBox.information(
                self, "Files Loaded",
                f"Loaded:\n• File A: {Path(excel_files[0]).name}\n• File B: {Path(excel_files[1]).name}"
            )
        elif len(excel_files) == 1:
            if self.file_a_path is None:
                self.file_a_display.setText(excel_files[0])
            else:
                self.file_b_display.setText(excel_files[0])
        else:
            QMessageBox.warning(
                self, "Invalid Files",
                "Please drop Excel files (.xlsx, .xls, .xlsm)"
            )

    # ---------- Key Column UI ----------
    def update_key_column_options(self, columns):
        # Clear existing
        while self.key_grid.count():
            item = self.key_grid.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        self.key_checkboxes.clear()

        cols_per_row = 4
        row = col = 0

        for name in columns:
            cb = QCheckBox(name)
            cb.setStyleSheet("font-size: 9pt; padding: 2px;")
            cb.toggled.connect(self.update_key_count)
            self.key_grid.addWidget(cb, row, col)
            self.key_checkboxes.append(cb)

            col += 1
            if col >= cols_per_row:
                col = 0
                row += 1
       
        # Update tiebreaker options (only for key-based mode)
        self.tiebreaker_combo.clear()
        self.tiebreaker_combo.addItem("(None - Optional)", None)
        for column in columns:
            self.tiebreaker_combo.addItem(column, column)
       
        self.update_key_count()

    def filter_key_columns(self, text):
        text = text.lower().strip()
        visible_count = 0
        for cb in self.key_checkboxes:
            visible = text in cb.text().lower()
            cb.setVisible(visible)
            if visible:
                visible_count += 1
       
        if text:
            self.key_count_label.setText(
                f"Showing {visible_count} of {len(self.key_checkboxes)} columns"
            )
        else:
            self.update_key_count()

    def toggle_all_keys(self, checked):
        for cb in self.key_checkboxes:
            if cb.isVisible():
                cb.setChecked(checked)

    def update_key_count(self):
        total = len(self.key_checkboxes)
        selected = sum(1 for cb in self.key_checkboxes if cb.isChecked())
        self.key_count_label.setText(
            f"Total: {total} columns | Selected: {selected}"
        )

    # ---------- Comparison ----------
    def run_comparison(self):
        keys = [cb.text() for cb in self.key_checkboxes if cb.isChecked()]
        if self.mode_key_based.isChecked():
            if not keys:
                QMessageBox.warning(
                    self, "Missing Keys",
                    "Please select at least one key column."
                )
                return
        else:
            keys = []  # No keys in position-based mode

        # Get tiebreaker column (only used in key-based mode with duplicate keys)
        tiebreaker = self.tiebreaker_combo.currentData()

        config = ComparisonConfig(
            key_columns=keys,
            alignment_method=AlignmentMethod.SECONDARY_SORT if tiebreaker else AlignmentMethod.POSITION,
            secondary_sort_column=tiebreaker,
            case_sensitive=self.case_sensitive.isChecked(),
            trim_whitespace=self.trim_whitespace.isChecked()
        )

        self.start_time = time.time()
        self.compare_btn.setEnabled(False)
        self.config_group.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)

        self.worker = ComparisonWorker(
            self.df_a, self.df_b, config,
            self.file_a_path, self.file_b_path
        )
        self.worker.progress.connect(self.statusBar().showMessage)
        self.worker.finished.connect(self.comparison_finished)
        self.worker.error.connect(self.comparison_error)
        self.worker.start()

    def comparison_finished(self, data):
        self.progress_bar.setVisible(False)
        self.compare_btn.setEnabled(True)
        self.config_group.setEnabled(True)

        elapsed = time.time() - self.start_time
        path = data["output_path"]
        result = data["result"]
        summary = result.summary
        metadata = result.comparison_metadata
       
        # Format time
        if elapsed < 60:
            time_str = f"{elapsed:.2f} seconds"
        else:
            minutes = int(elapsed // 60)
            seconds = elapsed % 60
            time_str = f"{minutes} min {seconds:.1f} sec"
       
        config = metadata.get('config')
       
        # Build configuration summary
        config_summary = ""
        if config:
            config_summary = f"""
⚙️ Comparison Configuration:
• Key Columns: {', '.join(config.key_columns) if config.key_columns else 'None (Position-based)'}
• Comparison Mode: {'Key-Based' if config.key_columns else 'Position-Based'}"""
            if config.secondary_sort_column:
                config_summary += f"\n• Tiebreaker Column: {config.secondary_sort_column}"
            config_summary += f"""
• Case Sensitive: {'Yes' if config.case_sensitive else 'No'}
• Trim Whitespace: {'Yes' if config.trim_whitespace else 'No'}
"""
       
        # Detailed results dialog
        msg = QMessageBox(self)
        msg.setWindowTitle("Comparison Complete")
        msg.setIcon(QMessageBox.Icon.Information)
        msg.setText(f"✅ Comparison completed in {time_str}!")
       
        details = f"""
📊 Summary Statistics:
• Total unique keys in File A: {summary['total_unique_keys_a']}
• Total unique keys in File B: {summary['total_unique_keys_b']}
• Keys in common: {summary['keys_in_common']}
• Keys only in File A: {summary['keys_only_in_a']}
• Keys only in File B: {summary['keys_only_in_b']}

📝 Row Comparison Results:
• Total rows compared: {summary['total_rows_compared']}
• ✅ Matching rows: {summary['match_count']}
• 🟡 Modified rows: {summary['modified_count']}
• 🟢 Added rows: {summary['added_row_count']}
• 🔴 Removed rows: {summary['removed_row_count']}
• 🔵 Rows in new keys: {summary['new_key_count']}
• 🟠 Rows in removed keys: {summary['removed_key_count']}
{config_summary}
📂 Report Location:
{path}

📁 Source Files:
• File A: {self.file_a_path}
  Sheet: {self.file_a_sheet}
• File B: {self.file_b_path}
  Sheet: {self.file_b_sheet}
"""
        msg.setDetailedText(details)
       
        open_btn = msg.addButton("📂 Open Report", QMessageBox.ButtonRole.AcceptRole)
        close_btn = msg.addButton("Close", QMessageBox.ButtonRole.RejectRole)
       
        msg.exec()
       
        if msg.clickedButton() == open_btn:
            if platform.system() == "Windows":
                os.startfile(path)
            elif platform.system() == "Darwin":
                os.system(f'open "{path}"')
            else:
                os.system(f'xdg-open "{path}"')
       
        self.statusBar().showMessage(f"✅ Comparison complete in {time_str}")

    def comparison_error(self, msg):
        self.progress_bar.setVisible(False)
        self.compare_btn.setEnabled(True)
        self.config_group.setEnabled(True)
        QMessageBox.critical(self, "Comparison Error", f"An error occurred:\n\n{msg}")
        self.statusBar().showMessage("❌ Comparison failed")

    # ---------- Settings ----------
    def load_settings(self):
        """Load saved settings"""
        geometry = self.settings.value("geometry")
        if geometry:
            self.restoreGeometry(geometry)
       
        self.case_sensitive.setChecked(
            self.settings.value("case_sensitive", False, type=bool)
        )
        self.trim_whitespace.setChecked(
            self.settings.value("trim_whitespace", True, type=bool)
        )

    def closeEvent(self, event):
        """Save settings on close"""
        self.settings.setValue("geometry", self.saveGeometry())
        self.settings.setValue("last_directory", self.last_directory)
        self.settings.setValue("case_sensitive", self.case_sensitive.isChecked())
        self.settings.setValue("trim_whitespace", self.trim_whitespace.isChecked())
        event.accept()


# =========================
# Entry Point
# =========================
def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = ExcelComparisonGUI()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()