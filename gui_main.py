"""
GUI for Excel Comparison Tool
Built with PySide6
"""

import sys
from pathlib import Path 
from datetime import datetime
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QGroupBox, QCheckBox,
    QComboBox, QProgressBar, QTextEdit, QMessageBox, QScrollArea,
    QGridLayout, QFrame, QLineEdit
)
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QFont, QIcon
import pandas as pd

from src.core import ComparisonEngine, ComparisonConfig, AlignmentMethod
from src.reports.report_generator import generate_comparison_report


class ComparisonWorker(QThread):
    """Worker thread for running comparison in background"""
   
    progress = Signal(str)  # Status messages
    finished = Signal(object)  # Comparison result
    error = Signal(str)  # Error messages
   
    def __init__(self, df_a, df_b, config, file_a_path, file_b_path):
        super().__init__()
        self.df_a = df_a
        self.df_b = df_b
        self.config = config
        self.file_a_path = file_a_path
        self.file_b_path = file_b_path
   
    def run(self):
        """Run comparison in background thread"""
        try:
            self.progress.emit("🔍 Comparing files...")
           
            engine = ComparisonEngine(self.config)
            result = engine.compare(self.df_a, self.df_b)
           
            self.progress.emit("📄 Generating Excel report...")
           
            # Generate report
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
           
            # Return result with output path
            result_data = {
                'result': result,
                'output_path': Path(output_file).resolve()
            }
           
            self.finished.emit(result_data)
           
        except Exception as e:
            self.error.emit(str(e))


class ExcelComparisonGUI(QMainWindow):
    """Main GUI window for Excel Comparison Tool"""
   
    def __init__(self):
        super().__init__()
       
        self.file_a_path = None
        self.file_b_path = None
        self.df_a = None
        self.df_b = None
        self.columns_a = []
        self.columns_b = []
        self.worker = None
        self.start_time = None
       
        self.init_ui()
   
    def init_ui(self):
        """Initialize the user interface"""
        self.setWindowTitle("Excel Comparison Tool v1.0")
        self.setMinimumSize(1000, 800)
        self.resize(1100, 850)
       
        # Central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
       
        # Main layout
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)
       
        # Title
        title = QLabel("Excel Comparison Tool")
        title_font = QFont()
        title_font.setPointSize(18)
        title_font.setBold(True)
        title.setFont(title_font)
        title.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title)
       
        # Subtitle
        subtitle = QLabel("Compare two Excel files with intelligent key-based matching")
        subtitle.setAlignment(Qt.AlignCenter)
        subtitle.setStyleSheet("color: gray; font-size: 11pt;")
        main_layout.addWidget(subtitle)
       
        # File selection section
        main_layout.addWidget(self._create_file_selection_section())
       
        # Configuration section (more space now)
        self.config_group = self._create_config_section()
        main_layout.addWidget(self.config_group)
        self.config_group.setEnabled(False)  # Disabled until files loaded
       
        # Compare button with integrated status
        main_layout.addWidget(self._create_compare_section())
       
        # Add stretch to push everything up
        main_layout.addStretch()
       
        # Status bar
        self.statusBar().showMessage("Ready - Select two Excel files to begin")
   
    def _create_file_selection_section(self):
        """Create file selection UI section"""
        group = QGroupBox("📁 1. Select Files")
        group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                font-size: 12pt;
                padding-top: 15px;
                margin-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
            }
        """)
        layout = QGridLayout()
        layout.setVerticalSpacing(15)
        layout.setHorizontalSpacing(10)
       
        # File A
        file_a_label = QLabel("File A:")
        file_a_label.setStyleSheet("font-weight: normal; font-size: 10pt;")
       
        self.file_a_display = QLineEdit()
        self.file_a_display.setPlaceholderText("No file selected")
        self.file_a_display.setReadOnly(True)
        self.file_a_display.setStyleSheet("""
            QLineEdit {
                padding: 8px;
                background-color: white;
                border: 1px solid #CCCCCC;
                border-radius: 3px;
                font-size: 10pt;
            }
        """)
       
        file_a_btn = QPushButton("Browse...")
        file_a_btn.setMinimumWidth(100)
        file_a_btn.setStyleSheet("""
            QPushButton {
                padding: 8px 15px;
                background-color: #F0F0F0;
                border: 1px solid #CCCCCC;
                border-radius: 3px;
                font-size: 10pt;
            }
            QPushButton:hover {
                background-color: #E0E0E0;
            }
        """)
        file_a_btn.clicked.connect(lambda: self.select_file('A'))
       
        layout.addWidget(file_a_label, 0, 0)
        layout.addWidget(self.file_a_display, 0, 1)
        layout.addWidget(file_a_btn, 0, 2)
       
        # Sheet selector for File A
        self.sheet_a_label = QLabel("Sheet:")
        self.sheet_a_label.setStyleSheet("font-weight: normal; font-size: 10pt;")
        self.sheet_a_combo = QComboBox()
        self.sheet_a_combo.setStyleSheet("font-size: 10pt; padding: 5px;")
        self.sheet_a_label.setVisible(False)
        self.sheet_a_combo.setVisible(False)
        layout.addWidget(self.sheet_a_label, 1, 0)
        layout.addWidget(self.sheet_a_combo, 1, 1, 1, 2)
       
        # File B
        file_b_label = QLabel("File B:")
        file_b_label.setStyleSheet("font-weight: normal; font-size: 10pt;")
       
        self.file_b_display = QLineEdit()
        self.file_b_display.setPlaceholderText("No file selected")
        self.file_b_display.setReadOnly(True)
        self.file_b_display.setStyleSheet("""
            QLineEdit {
                padding: 8px;
                background-color: white;
                border: 1px solid #CCCCCC;
                border-radius: 3px;
                font-size: 10pt;
            }
        """)
       
        file_b_btn = QPushButton("Browse...")
        file_b_btn.setMinimumWidth(100)
        file_b_btn.setStyleSheet("""
            QPushButton {
                padding: 8px 15px;
                background-color: #F0F0F0;
                border: 1px solid #CCCCCC;
                border-radius: 3px;
                font-size: 10pt;
            }
            QPushButton:hover {
                background-color: #E0E0E0;
            }
        """)
        file_b_btn.clicked.connect(lambda: self.select_file('B'))
       
        layout.addWidget(file_b_label, 2, 0)
        layout.addWidget(self.file_b_display, 2, 1)
        layout.addWidget(file_b_btn, 2, 2)
       
        # Sheet selector for File B
        self.sheet_b_label = QLabel("Sheet:")
        self.sheet_b_label.setStyleSheet("font-weight: normal; font-size: 10pt;")
        self.sheet_b_combo = QComboBox()
        self.sheet_b_combo.setStyleSheet("font-size: 10pt; padding: 5px;")
        self.sheet_b_label.setVisible(False)
        self.sheet_b_combo.setVisible(False)
        layout.addWidget(self.sheet_b_label, 3, 0)
        layout.addWidget(self.sheet_b_combo, 3, 1, 1, 2)
       
        layout.setColumnStretch(1, 1)
        group.setLayout(layout)
        return group
   
    def _create_config_section(self):
        """Create configuration UI section"""
        group = QGroupBox("⚙️ 2. Configure Comparison")
        group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                font-size: 12pt;
                padding-top: 15px;
                margin-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
            }
        """)
        layout = QVBoxLayout()
        layout.setSpacing(15)
       
        # Key columns selection
        key_label = QLabel("Select key columns to uniquely identify rows:")
        key_label.setStyleSheet("font-weight: normal; margin-top: 5px; font-size: 10pt;")
        layout.addWidget(key_label)
       
        # Key checkboxes in a grid layout (2 columns)
        self.key_checkboxes_widget = QWidget()
        self.key_checkboxes_layout = QGridLayout(self.key_checkboxes_widget)
        self.key_checkboxes_layout.setSpacing(10)
        self.key_checkboxes = []
       
        # Wrap in scroll area
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setMinimumHeight(100)
        scroll.setMaximumHeight(150)
        scroll.setStyleSheet("""
            QScrollArea {
                border: 1px solid #CCCCCC;
                border-radius: 3px;
                background-color: white;
            }
            QCheckBox {
                padding: 3px;
                font-size: 10pt;
            }
        """)
        scroll.setWidget(self.key_checkboxes_widget)
        layout.addWidget(scroll)
       
        # Comparison Options
        options_label = QLabel("Comparison Options:")
        options_label.setStyleSheet("font-weight: normal; margin-top: 10px; font-size: 10pt;")
        layout.addWidget(options_label)
       
        options_layout = QGridLayout()
        options_layout.setVerticalSpacing(12)
        options_layout.setHorizontalSpacing(15)
       
        # Alignment method
        alignment_label = QLabel("Alignment Method:")
        alignment_label.setStyleSheet("font-size: 10pt; font-weight: normal;")
        options_layout.addWidget(alignment_label, 0, 0, Qt.AlignLeft | Qt.AlignVCenter)
       
        self.alignment_combo = QComboBox()
        self.alignment_combo.setMinimumHeight(32)
        self.alignment_combo.setStyleSheet("""
            QComboBox {
                font-size: 10pt;
                padding: 5px;
                border: 1px solid #CCCCCC;
                border-radius: 3px;
            }
        """)
        self.alignment_combo.addItem("Position-based (1st row → 1st row)", AlignmentMethod.POSITION)
        self.alignment_combo.addItem("Secondary Sort Column", AlignmentMethod.SECONDARY_SORT)
        self.alignment_combo.currentIndexChanged.connect(self.on_alignment_changed)
        options_layout.addWidget(self.alignment_combo, 0, 1, 1, 2)
       
        # Secondary sort column (initially hidden)
        self.secondary_sort_label = QLabel("Secondary Sort Column:")
        self.secondary_sort_label.setStyleSheet("font-size: 10pt; font-weight: normal;")
        self.secondary_sort_combo = QComboBox()
        self.secondary_sort_combo.setMinimumHeight(32)
        self.secondary_sort_combo.setStyleSheet("""
            QComboBox {
                font-size: 10pt;
                padding: 5px;
                border: 1px solid #CCCCCC;
                border-radius: 3px;
            }
        """)
        self.secondary_sort_label.setVisible(False)
        self.secondary_sort_combo.setVisible(False)
        options_layout.addWidget(self.secondary_sort_label, 1, 0, Qt.AlignLeft | Qt.AlignVCenter)
        options_layout.addWidget(self.secondary_sort_combo, 1, 1, 1, 2)
       
        # Checkboxes with better styling
        self.case_sensitive_check = QCheckBox("Case Sensitive Comparison")
        self.case_sensitive_check.setStyleSheet("font-size: 10pt; padding: 5px;")
        options_layout.addWidget(self.case_sensitive_check, 2, 1, 1, 2)
       
        self.trim_whitespace_check = QCheckBox("Trim Whitespace")
        self.trim_whitespace_check.setChecked(True)
        self.trim_whitespace_check.setStyleSheet("font-size: 10pt; padding: 5px;")
        options_layout.addWidget(self.trim_whitespace_check, 3, 1, 1, 2)
       
        options_layout.setColumnStretch(1, 1)
        layout.addLayout(options_layout)
       
        group.setLayout(layout)
        return group
   
    def _create_progress_section(self):
        """Create progress UI section"""
        group = QGroupBox("📊 Status")
        group.setStyleSheet("QGroupBox { font-weight: bold; font-size: 11pt; }")
        layout = QVBoxLayout()
       
        self.progress_label = QLabel("Ready to compare")
        self.progress_label.setStyleSheet("padding: 5px;")
        layout.addWidget(self.progress_label)
       
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setMinimumHeight(25)
        layout.addWidget(self.progress_bar)
       
        group.setLayout(layout)
        return group
   
    def _create_compare_section(self):
        """Create compare button with integrated status"""
        group = QGroupBox("🔍 3. Start Comparison")
        group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                font-size: 12pt;
                padding-top: 15px;
                margin-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
            }
        """)
        layout = QVBoxLayout()
        layout.setSpacing(10)
       
        # Compare button
        self.compare_btn = QPushButton("🔍 Compare Files")
        self.compare_btn.setMinimumHeight(50)
        self.compare_btn.setStyleSheet("""
            QPushButton {
                background-color: #5B7FB8;
                color: white;
                font-size: 13pt;
                font-weight: bold;
                border-radius: 5px;
                border: none;
            }
            QPushButton:hover {
                background-color: #4A6A9E;
            }
            QPushButton:disabled {
                background-color: #CCCCCC;
                color: #888888;
            }
        """)
        self.compare_btn.clicked.connect(self.run_comparison)
        self.compare_btn.setEnabled(False)
        layout.addWidget(self.compare_btn)
       
        # Progress bar (hidden by default)
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setMinimumHeight(25)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid #CCCCCC;
                border-radius: 3px;
                text-align: center;
            }
            QProgressBar::chunk {
                background-color: #5B7FB8;
            }
        """)
        layout.addWidget(self.progress_bar)
       
        group.setLayout(layout)
       
        # Status section (separate, below the compare section)
        status_widget = QWidget()
        status_layout = QVBoxLayout(status_widget)
        status_layout.setContentsMargins(0, 10, 0, 0)
       
        self.progress_label = QLabel("")
        self.progress_label.setAlignment(Qt.AlignCenter)
        self.progress_label.setStyleSheet("""
            padding: 10px;
            background-color: #D4EDDA;
            border: 1px solid #C3E6CB;
            border-radius: 3px;
            font-size: 10pt;
            color: #155724;
        """)
        self.progress_label.setVisible(False)
        status_layout.addWidget(self.progress_label)
       
        # Create container
        container = QWidget()
        container_layout = QVBoxLayout(container)
        container_layout.setSpacing(5)
        container_layout.addWidget(group)
        container_layout.addWidget(status_widget)
       
        return container
   

    def select_file(self, file_type):
        """Handle file selection"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            f"Select Excel File {file_type}",
            "",
            "Excel Files (*.xlsx *.xls *.xlsm)"
        )
       
        if not file_path:
            return
       
        try:
            # Get sheet names
            excel_file = pd.ExcelFile(file_path)
            sheet_names = excel_file.sheet_names
           
            # Shorten path for display (keep first and last parts)
            display_path = self._shorten_path(file_path)
           
            if file_type == 'A':
                self.file_a_path = file_path
                self.file_a_display.setText(file_path)
                self.file_a_display.setToolTip(file_path)  # Full path on hover
               
                # Disconnect old signal if exists
                try:
                    self.sheet_a_combo.currentTextChanged.disconnect()
                except:
                    pass
               
                # Show sheet selector if multiple sheets
                self.sheet_a_combo.clear()
                for sheet in sheet_names:
                    self.sheet_a_combo.addItem(sheet)
               
                if len(sheet_names) > 1:
                    self.sheet_a_label.setVisible(True)
                    self.sheet_a_combo.setVisible(True)
                else:
                    self.sheet_a_label.setVisible(False)
                    self.sheet_a_combo.setVisible(False)
               
                # Load default sheet
                self.df_a = pd.read_excel(file_path, sheet_name=sheet_names[0])
                self.columns_a = list(self.df_a.columns)
               
                # Connect sheet change AFTER setting initial value
                self.sheet_a_combo.currentTextChanged.connect(lambda: self._reload_sheet('A'))
               
                self.statusBar().showMessage(
                    f"File A loaded: {len(self.df_a)} rows, {len(self.df_a.columns)} columns"
                )
            else:
                self.file_b_path = file_path
                self.file_b_display.setText(file_path)
                self.file_b_display.setToolTip(file_path)  # Full path on hover
               
                # Disconnect old signal if exists
                try:
                    self.sheet_b_combo.currentTextChanged.disconnect()
                except:
                    pass
               
                # Show sheet selector if multiple sheets
                self.sheet_b_combo.clear()
                for sheet in sheet_names:
                    self.sheet_b_combo.addItem(sheet)
               
                if len(sheet_names) > 1:
                    self.sheet_b_label.setVisible(True)
                    self.sheet_b_combo.setVisible(True)
                else:
                    self.sheet_b_label.setVisible(False)
                    self.sheet_b_combo.setVisible(False)
               
                # Load default sheet
                self.df_b = pd.read_excel(file_path, sheet_name=sheet_names[0])
                self.columns_b = list(self.df_b.columns)
               
                # Connect sheet change AFTER setting initial value
                self.sheet_b_combo.currentTextChanged.connect(lambda: self._reload_sheet('B'))
               
                self.statusBar().showMessage(
                    f"File B loaded: {len(self.df_b)} rows, {len(self.df_b.columns)} columns"
                )
           
            # Check if both files are loaded
            if self.df_a is not None and self.df_b is not None:
                self.update_key_column_options()
                self.config_group.setEnabled(True)
                self.compare_btn.setEnabled(True)
               
        except Exception as e:
            QMessageBox.critical(self, "Error Loading File", f"Could not load file:\n{str(e)}")
   
    def _shorten_path(self, file_path, max_length=60):
        """Shorten path for display while keeping beginning and end"""
        if len(file_path) <= max_length:
            return file_path
       
        path_obj = Path(file_path)
        filename = path_obj.name
        parent_parts = list(path_obj.parent.parts)
       
        # Start with drive and filename
        if len(parent_parts) > 0:
            shortened = parent_parts[0] + "\\"
        else:
            shortened = ""
       
        # Add as many middle parts as fit
        available_length = max_length - len(shortened) - len(filename) - 6  # 6 for "...\"
       
        if available_length < 0:
            # Just show drive...filename
            return f"{parent_parts[0]}\\...\\{filename}"
       
        # Try to fit some middle directories
        middle_parts = parent_parts[1:]
        included_parts = []
        current_length = 0
       
        for part in reversed(middle_parts):
            if current_length + len(part) + 1 <= available_length:
                included_parts.insert(0, part)
                current_length += len(part) + 1
            else:
                break
       
        if included_parts:
            shortened += "...\\" + "\\".join(included_parts) + "\\" + filename
        else:
            shortened += "...\\" + filename
       
        return shortened
   
    def _reload_sheet(self, file_type):
        """Reload data when sheet selection changes"""
        try:
            if file_type == 'A':
                sheet_name = self.sheet_a_combo.currentText()
                if not sheet_name:  # Empty selection, skip
                    return
                self.df_a = pd.read_excel(self.file_a_path, sheet_name=sheet_name)
                self.columns_a = list(self.df_a.columns)
                self.statusBar().showMessage(
                    f"File A sheet '{sheet_name}': {len(self.df_a)} rows, {len(self.df_a.columns)} columns"
                )
            else:
                sheet_name = self.sheet_b_combo.currentText()
                if not sheet_name:  # Empty selection, skip
                    return
                self.df_b = pd.read_excel(self.file_b_path, sheet_name=sheet_name)
                self.columns_b = list(self.df_b.columns)
                self.statusBar().showMessage(
                    f"File B sheet '{sheet_name}': {len(self.df_b)} rows, {len(self.df_b.columns)} columns"
                )
           
            # Update key column options
            if self.df_a is not None and self.df_b is not None:
                self.update_key_column_options()
               
        except Exception as e:
            QMessageBox.critical(self, "Error Loading Sheet", f"Could not load sheet:\n{str(e)}")
   
    def update_key_column_options(self):
        """Update key column checkboxes based on loaded files"""
        # Clear existing checkboxes
        for checkbox in self.key_checkboxes:
            checkbox.deleteLater()
        self.key_checkboxes.clear()
       
        # Clear layout
        for i in reversed(range(self.key_checkboxes_layout.count())):
            self.key_checkboxes_layout.itemAt(i).widget().setParent(None)
       
        # Find common columns
        common_columns = set(self.columns_a) & set(self.columns_b)
       
        if not common_columns:
            QMessageBox.warning(
                self,
                "No Common Columns",
                "Files have no columns in common. Cannot perform key-based comparison."
            )
            return
       
        # Create checkboxes for common columns in a 2-column grid
        sorted_columns = sorted(common_columns)
        row = 0
        col = 0
        for column_name in sorted_columns:
            checkbox = QCheckBox(column_name)
            checkbox.setStyleSheet("font-size: 10pt;")
            self.key_checkboxes.append(checkbox)
            self.key_checkboxes_layout.addWidget(checkbox, row, col)
           
            col += 1
            if col > 1:  # 2 columns
                col = 0
                row += 1
       
        # Update secondary sort options
        self.secondary_sort_combo.clear()
        self.secondary_sort_combo.addItem("(None)", None)
        for col in sorted_columns:
            self.secondary_sort_combo.addItem(col, col)
   
    def on_alignment_changed(self):
        """Handle alignment method change"""
        method = self.alignment_combo.currentData()
        show_secondary = (method == AlignmentMethod.SECONDARY_SORT)
        self.secondary_sort_label.setVisible(show_secondary)
        self.secondary_sort_combo.setVisible(show_secondary)
   
    def run_comparison(self):
        """Start comparison process"""
        # Validate key selection
        selected_keys = [cb.text() for cb in self.key_checkboxes if cb.isChecked()]
       
        if not selected_keys:
            QMessageBox.warning(
                self,
                "No Keys Selected",
                "Please select at least one key column for comparison."
            )
            return
       
        # Build configuration
        alignment_method = self.alignment_combo.currentData()
        secondary_sort = None
        if alignment_method == AlignmentMethod.SECONDARY_SORT:
            secondary_sort = self.secondary_sort_combo.currentData()
            if not secondary_sort:
                QMessageBox.warning(
                    self,
                    "No Secondary Sort Selected",
                    "Please select a secondary sort column or use position-based alignment."
                )
                return
       
        config = ComparisonConfig(
            key_columns=selected_keys,
            alignment_method=alignment_method,
            secondary_sort_column=secondary_sort,
            case_sensitive=self.case_sensitive_check.isChecked(),
            trim_whitespace=self.trim_whitespace_check.isChecked()
        )
       
        # Record start time
        import time
        self.start_time = time.time()
       
        # Disable UI during comparison
        self.compare_btn.setEnabled(False)
        self.config_group.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)  # Indeterminate
       
        # Start worker thread
        self.worker = ComparisonWorker(
            self.df_a, self.df_b, config,
            self.file_a_path, self.file_b_path
        )
        self.worker.progress.connect(self.update_progress)
        self.worker.finished.connect(self.comparison_finished)
        self.worker.error.connect(self.comparison_error)
        self.worker.start()
   
    def update_progress(self, message):
        """Update progress display"""
        self.progress_label.setText(message)
        self.statusBar().showMessage(message)
   
    def comparison_finished(self, result_data):
        """Handle comparison completion"""
        import time
        elapsed_time = time.time() - self.start_time
       
        result = result_data['result']
        output_path = result_data['output_path']
       
        # Re-enable UI
        self.compare_btn.setEnabled(True)
        self.config_group.setEnabled(True)
        self.progress_bar.setVisible(False)
       
        # Format elapsed time
        if elapsed_time < 60:
            time_str = f"{elapsed_time:.2f} seconds"
        else:
            minutes = int(elapsed_time // 60)
            seconds = elapsed_time % 60
            time_str = f"{minutes} min {seconds:.1f} sec"
       
        # Build detailed results message
        summary = result.summary
        results_message = f"""Comparison Complete in {time_str}!

Keys:
• Total unique keys in File A: {summary['total_unique_keys_a']}
• Total unique keys in File B: {summary['total_unique_keys_b']}
• Keys in common: {summary['keys_in_common']}
• Keys only in A: {summary['keys_only_in_a']}
• Keys only in B: {summary['keys_only_in_b']}

Rows:
• Total rows compared: {summary['total_rows_compared']}
• ✅ Matching rows: {summary['match_count']}
• 🟡 Modified rows: {summary['modified_count']}
• 🟢 Added rows: {summary['added_row_count']}
• 🔴 Removed rows: {summary['removed_row_count']}
• 🔵 Rows in new keys: {summary['new_key_count']}
• 🟠 Rows in removed keys: {summary['removed_key_count']}

Report Location:
{output_path}
"""
       
        self.report_path = output_path
        self.progress_label.setText(f"✅ Status: Comparison complete in {time_str} - Report generated")
        self.progress_label.setVisible(True)
        self.statusBar().showMessage(f"Comparison complete in {time_str} - Report generated")
       
        # Show detailed results dialog with option to open report
        msg = QMessageBox(self)
        msg.setWindowTitle("Comparison Complete")
        msg.setText(f"✅ Comparison completed successfully in {time_str}!")
        msg.setDetailedText(results_message)
        msg.setIcon(QMessageBox.Information)
       
        # Add custom buttons
        open_btn = msg.addButton("📂 Open Report", QMessageBox.AcceptRole)
        close_btn = msg.addButton("Close", QMessageBox.RejectRole)
       
        msg.exec()
       
        # Check which button was clicked
        if msg.clickedButton() == open_btn:
            self.open_report()
   
    def comparison_error(self, error_msg):
        """Handle comparison error"""
        # Re-enable UI
        self.compare_btn.setEnabled(True)
        self.config_group.setEnabled(True)
        self.progress_bar.setVisible(False)
       
        self.progress_label.setText("❌ Comparison failed")
        self.statusBar().showMessage("Comparison failed - See error details")
       
        QMessageBox.critical(
            self,
            "Comparison Error",
            f"An error occurred during comparison:\n\n{error_msg}"
        )
   
    def open_report(self):
        """Open the generated Excel report"""
        if hasattr(self, 'report_path'):
            import os
            import platform
           
            try:
                if platform.system() == 'Windows':
                    os.startfile(self.report_path)
                elif platform.system() == 'Darwin':  # macOS
                    os.system(f'open "{self.report_path}"')
                else:  # Linux
                    os.system(f'xdg-open "{self.report_path}"')
            except Exception as e:
                QMessageBox.warning(
                    self,
                    "Cannot Open File",
                    f"Could not open report automatically:\n{str(e)}\n\nPlease open manually:\n{self.report_path}"
                )


def main():
    """Main entry point for GUI"""
    app = QApplication(sys.argv)
    app.setStyle('Fusion')  # Modern look
   
    window = ExcelComparisonGUI()
    window.show()
   
    sys.exit(app.exec())


if __name__ == "__main__":
    main()