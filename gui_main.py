"""
GUI for Excel Comparison Tool
Built with PySide6
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
    QComboBox, QProgressBar, QMessageBox, QScrollArea,
    QGridLayout, QLineEdit
)
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QFont

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
        self.df_a = None
        self.df_b = None
        self.key_checkboxes = []
        self.worker = None
        self.start_time = None
        self.init_ui()

    # ---------- UI ----------
    def init_ui(self):
        self.setWindowTitle("Excel Comparison Tool v1.0")
        self.resize(1100, 850)

        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)

        title = QLabel("Excel Comparison Tool")
        # title.setFont(QFont("", 18, QFont.Bold))
        title.setFont(self.ui_font(size=18, bold=True))
        title.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title)

        subtitle = QLabel("Compare two Excel files using key-based matching")
        subtitle.setAlignment(Qt.AlignCenter)
        subtitle.setStyleSheet("color: gray;")
        main_layout.addWidget(subtitle)

        main_layout.addWidget(self.create_file_section())
        main_layout.addWidget(self.create_config_section())
        main_layout.addWidget(self.create_compare_section())
        main_layout.addStretch()

        self.statusBar().showMessage("Ready")

    def ui_font(self, size=10, bold=False):
        font = QFont()
        font.setPointSize(size)
        if bold:
            font.setWeight(QFont.Bold)
        return font

    # ---------- File Section ----------
    def create_file_section(self):
        group = QGroupBox("1. Select Files")
        layout = QGridLayout(group)
        layout.setSpacing(10)

        self.file_a_display = QLineEdit()
        self.file_a_display.setReadOnly(True)
        btn_a = QPushButton("Browse A")
        btn_a.clicked.connect(lambda: self.select_file("A"))

        self.file_b_display = QLineEdit()
        self.file_b_display.setReadOnly(True)
        btn_b = QPushButton("Browse B")
        btn_b.clicked.connect(lambda: self.select_file("B"))

        layout.addWidget(QLabel("File A:"), 0, 0)
        layout.addWidget(self.file_a_display, 0, 1)
        layout.addWidget(btn_a, 0, 2)

        layout.addWidget(QLabel("File B:"), 1, 0)
        layout.addWidget(self.file_b_display, 1, 1)
        layout.addWidget(btn_b, 1, 2)

        layout.setColumnStretch(1, 1)
        return group

    # ---------- Config Section (REPLACED) ----------
    def create_config_section(self):
        self.config_group = QGroupBox("2. Configure Comparison")
        self.config_group.setEnabled(False)
        layout = QVBoxLayout(self.config_group)
        layout.setSpacing(12)

        # ---- Key Columns ----
        key_group = QGroupBox("Select Key Columns (unique row identifier)")
        key_layout = QVBoxLayout(key_group)

        self.key_filter = QLineEdit()
        self.key_filter.setPlaceholderText("Filter columns...")
        self.key_filter.textChanged.connect(self.filter_key_columns)
        key_layout.addWidget(self.key_filter)

        self.key_scroll = QScrollArea()
        self.key_scroll.setWidgetResizable(True)
        self.key_scroll.setMaximumHeight(220)

        self.key_container = QWidget()
        self.key_grid = QGridLayout(self.key_container)
        self.key_grid.setSpacing(8)
        self.key_grid.setContentsMargins(6, 6, 6, 6)

        self.key_scroll.setWidget(self.key_container)
        key_layout.addWidget(self.key_scroll)
        layout.addWidget(key_group)

        # ---- Options ----
        self.case_sensitive = QCheckBox("Case Sensitive Comparison")
        self.trim_whitespace = QCheckBox("Trim Whitespace")
        self.trim_whitespace.setChecked(True)

        layout.addWidget(self.case_sensitive)
        layout.addWidget(self.trim_whitespace)

        return self.config_group

    # ---------- Compare Section ----------
    def create_compare_section(self):
        group = QGroupBox("3. Start Comparison")
        layout = QVBoxLayout(group)

        self.compare_btn = QPushButton("Compare Files")
        self.compare_btn.setMinimumHeight(45)
        self.compare_btn.setEnabled(False)
        self.compare_btn.clicked.connect(self.run_comparison)

        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)

        layout.addWidget(self.compare_btn)
        layout.addWidget(self.progress_bar)
        return group

    # ---------- File Handling ----------
    def select_file(self, which):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select Excel File", "", "Excel Files (*.xlsx *.xls *.xlsm)"
        )
        if not path:
            return

        df = pd.read_excel(path)
        if which == "A":
            self.file_a_path = path
            self.df_a = df
            self.file_a_display.setText(path)
        else:
            self.file_b_path = path
            self.df_b = df
            self.file_b_display.setText(path)

        if self.df_a is not None and self.df_b is not None:
            common_cols = [col for col in self.df_a.columns if col in self.df_b.columns]
            self.update_key_column_options(common_cols)
            self.config_group.setEnabled(True)
            self.compare_btn.setEnabled(True)

    # ---------- Key Column UI ----------
    def update_key_column_options(self, columns):
        while self.key_grid.count():
            item = self.key_grid.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        self.key_checkboxes.clear()

        cols_per_row = 3
        row = col = 0

        for name in columns:
            cb = QCheckBox(name)
            self.key_grid.addWidget(cb, row, col)
            self.key_checkboxes.append(cb)

            col += 1
            if col >= cols_per_row:
                col = 0
                row += 1

    def filter_key_columns(self, text):
        text = text.lower().strip()
        for cb in self.key_checkboxes:
            cb.setVisible(text in cb.text().lower())

    # ---------- Comparison ----------
    def run_comparison(self):
        keys = [cb.text() for cb in self.key_checkboxes if cb.isChecked()]
        if not keys:
            QMessageBox.warning(self, "Missing Keys", "Select at least one key column.")
            return

        config = ComparisonConfig(
            key_columns=keys,
            alignment_method=AlignmentMethod.POSITION,
            secondary_sort_column=None,
            case_sensitive=self.case_sensitive.isChecked(),
            trim_whitespace=self.trim_whitespace.isChecked()
        )

        self.start_time = time.time()
        self.compare_btn.setEnabled(False)
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

        elapsed = time.time() - self.start_time
        path = data["output_path"]

        QMessageBox.information(
            self,
            "Done",
            f"Comparison completed in {elapsed:.2f}s\n\nReport:\n{path}"
        )

        if platform.system() == "Windows":
            os.startfile(path)

    def comparison_error(self, msg):
        self.progress_bar.setVisible(False)
        self.compare_btn.setEnabled(True)
        QMessageBox.critical(self, "Error", msg)


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
