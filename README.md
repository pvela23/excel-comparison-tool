# excel-comparison-tool

A professional desktop application for comparing Excel files with intelligent key-based matching. Built for analysts, accountants, and data professionals who need to reconcile complex Excel files quickly and accurately.

![Version](https://img.shields.io/badge/version-1.0.0-blue)
![Python](https://img.shields.io/badge/python-3.8+-green)
![License](https://img.shields.io/badge/license-MIT-blue)

## 🎯 Overview

Excel Comparison Tool solves a common problem: comparing two Excel files when rows don't match one-to-one. Unlike basic diff tools, it intelligently matches rows using key columns you specify, handles multiple rows per key, and produces professional, color-coded Excel reports.

### Key Features

✅ **Intelligent Key-Based Matching** - Match rows by Policy Number, ID, or any combination of columns  
✅ **Multi-Row Support** - Handles cases where one key has multiple rows (e.g., one policy with multiple coverages)  
✅ **Flexible Alignment** - Position-based or sort by secondary column before matching  
✅ **Non-Destructive** - Original files never modified, comparison output to new Excel file  
✅ **Professional Reports** - Color-coded Excel output with Summary, Aligned Diff, and Legend sheets  
✅ **User-Friendly GUI** - Drag & drop files, filter columns, keyboard shortcuts  
✅ **Sheet Selection** - Compare specific sheets from multi-sheet workbooks  
✅ **Large File Support** - Handles files with 500K+ rows with memory warnings  

---

## 📸 Screenshots

### Main Interface
```
┌─────────────────────────────────────────────────────────┐
│                Excel Comparison Tool                    │
│      Compare two Excel files using key-based matching  │
├─────────────────────────────────────────────────────────┤
│ 📁 1. Select Files                                      │
│   File A: [C:\data\report_2024.xlsx]      [Browse...]  │
│   File B: [C:\data\report_2025.xlsx]      [Browse...]  │
├─────────────────────────────────────────────────────────┤
│ ⚙️ 2. Configure Comparison                             │
│   Select Key Columns:                                   │
│   ☑ Policy Number  ☑ Effective Date  ☐ Coverage       │
│                                                         │
│   Alignment Method: [Position-based ▼]                 │
│   ☐ Case Sensitive  ☑ Trim Whitespace                  │
├─────────────────────────────────────────────────────────┤
│ 🔍 3. Start Comparison                                  │
│               [Compare Files]                           │
└─────────────────────────────────────────────────────────┘
```

### Excel Report Output
- **Summary Sheet**: Key statistics and row counts
- **Aligned Diff Sheet**: Side-by-side comparison with color coding
  - 🟢 Green = Added rows
  - 🔴 Red = Removed rows
  - 🟡 Yellow = Modified cells
  - ⚪ White = Unchanged
- **Legend Sheet**: Color meanings and configuration used

---

## 🚀 Quick Start

### Installation

#### Prerequisites
- Python 3.8 or higher
- Windows, macOS, or Linux

#### Install from Source

```bash
# Clone the repository
git clone https://github.com/yourusername/excel-comparison-tool.git
cd excel-comparison-tool

# Create virtual environment
python -m venv venv

# Activate virtual environment
# Windows:
venv\Scripts\activate
# macOS/Linux:
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt
```

### Running the Tool

```bash
# GUI Mode (recommended)
python gui_main.py

# CLI Mode (for automation)
python main.py
```

---

## 📖 Usage Guide

### Basic Workflow

1. **Select Files**
   - Click "Browse..." or drag & drop Excel files
   - If file has multiple sheets, select which to compare
   - Files can be in same folder or different locations

2. **Configure Comparison**
   - **Select Key Columns**: Check boxes for columns that uniquely identify rows
   - **Choose Alignment Method**:
     - *Position-based*: 1st row in File A → 1st row in File B
     - *Secondary Sort*: Sort by column (e.g., Invoice Number) before matching
   - **Set Options**:
     - Case Sensitive: Match "ABC" vs "abc" as different
     - Trim Whitespace: Ignore leading/trailing spaces

3. **Compare**
   - Click "Compare Files" or press `Ctrl+Enter`
   - Progress bar shows status
   - Report auto-opens when complete

### Example Use Cases

#### Insurance Policy Reconciliation
```
Scenario: Compare current vs. updated policy data
Key Columns: Policy Number, Effective Date
Alignment: Position-based
Result: Identifies premium changes, new coverages, cancelled policies
```

#### Invoice Audit
```
Scenario: Verify invoice system export vs. manual corrections
Key Columns: Invoice Number, Line Number
Alignment: Secondary Sort by Line Number
Result: Shows which line items were modified
```

#### Claims Comparison
```
Scenario: Compare claims data before/after adjustment
Key Columns: Claim ID, Coverage Type
Alignment: Position-based
Result: Tracks claim amount changes by coverage
```

---

## 🎨 Features in Detail

### Intelligent Row Matching

**Problem**: Files have different row counts
```
File A:                    File B:
Policy 12345 - Auto        Policy 12345 - Auto    (modified)
Policy 12345 - Home        Policy 12345 - Home    (same)
                           Policy 12345 - Life    (new)
Policy 67890 - Auto        Policy 67890 - Auto    (same)
```

**Solution**: Tool groups by key (Policy 12345), then compares within group
- 1 modified row (Auto premium changed)
- 1 matching row (Home unchanged)
- 1 added row (Life is new)

### Column Filtering & Selection

- **Filter Box**: Type "date" to show only date columns
- **Select All/Deselect All**: Quick selection for many columns
- **Real-time Count**: "Total: 15 columns | Selected: 3"

### Drag & Drop Support

- Drop 1 file → Loads into first empty slot
- Drop 2 files → Loads both (File A and File B)
- Supports multi-file selection

### Keyboard Shortcuts

- `Ctrl+Enter` - Start comparison
- Works anywhere in the application

### Settings Persistence

Settings saved between sessions:
- Last directory browsed
- Window size and position
- Case sensitivity preference
- Trim whitespace preference

---

## 📋 Requirements

### Python Dependencies

```txt
pandas>=2.0.0
openpyxl>=3.1.0
numpy>=1.24.0
PySide6>=6.5.0
```

### System Requirements

- **RAM**: 4GB minimum, 8GB recommended for large files
- **Disk**: 100MB for application, additional space for reports
- **Display**: 1024x768 minimum resolution

### File Support

- **Formats**: .xlsx, .xls, .xlsm
- **Size**: Up to 1M rows (500K+ shows warning)
- **Sheets**: Multi-sheet workbooks supported

---

## ⚙️ Configuration

### ComparisonConfig Object

```python
from src.core import ComparisonConfig, AlignmentMethod

config = ComparisonConfig(
    key_columns=[],           # Required
    alignment_method=AlignmentMethod.POSITION,  # POSITION or SECONDARY_SORT
    secondary_sort_column='',  # Optional, for SECONDARY_SORT
    case_sensitive=False,                    # Default: False
    trim_whitespace=True                     # Default: True
)
```

### Environment Variables

```bash
# Optional: Set default comparison directory
export EXCEL_COMP_DEFAULT_DIR="/path/to/files"

# Optional: Disable auto-open of reports
export EXCEL_COMP_NO_AUTOOPEN=1
```

---

## 🔧 Advanced Usage

### CLI Mode (Automation)

```python
# main.py with custom paths
from src.core import ComparisonEngine, ComparisonConfig, AlignmentMethod
from src.reports.report_generator import generate_comparison_report
import pandas as pd

# Load files
df_a = pd.read_excel('file_a.xlsx', dtype=str)
df_b = pd.read_excel('file_b.xlsx', dtype=str)

# Configure
config = ComparisonConfig(
    key_columns=['Policy', 'EFF'],
    alignment_method=AlignmentMethod.POSITION
)

# Compare
engine = ComparisonEngine(config)
result = engine.compare(df_a, df_b)

# Generate report
generate_comparison_report(
    output_path='report.xlsx',
    summary=result.summary,
    aligned_data=result.aligned_data,
    metadata=result.comparison_metadata,
    file_a_path='file_a.xlsx',
    file_b_path='file_b.xlsx'
)

print(f"Modified: {result.summary['modified_count']}")
print(f"Added: {result.summary['added_row_count']}")
```

### Batch Processing

```python
import os
from pathlib import Path

# Compare all files in directory
data_dir = Path('data')
for file in data_dir.glob('report_*.xlsx'):
    file_a = file
    file_b = file.with_name(file.stem + '_updated.xlsx')
    
    if file_b.exists():
        # Run comparison...
        print(f"Comparing {file.name} vs {file_b.name}")
```

---

## 🐛 Troubleshooting

### Common Issues

#### "No module named 'pandas'"
```bash
# Make sure virtual environment is activated
venv\Scripts\activate  # Windows
source venv/bin/activate  # macOS/Linux

# Install dependencies
pip install -r requirements.txt
```

#### "Cannot access file (may be open in Excel)"
- Close the Excel file before loading
- Check file isn't locked by another process

#### "No common columns found"
- Files must have at least one column with same name
- Column names are case-sensitive
- Use column mapping if names differ

#### Large file loads slowly
- This is expected for 500K+ rows
- Loading as dtype=str takes longer but prevents errors
- Consider comparing smaller subsets first

#### Report doesn't auto-open
- Windows: Check file associations for .xlsx
- macOS: Ensure Excel or compatible app is default
- Manual: Report saved in tool directory

---

## 🏗️ Project Structure

```
excel-comparison-tool/
├── src/
│   ├── core/
│   │   ├── __init__.py
│   │   └── comparison_engine.py     # Core comparison logic
│   ├── reports/
│   │   ├── __init__.py
│   │   └── report_generator.py      # Excel report generation
│   └── ui/
│       └── (future GUI components)
├── tests/
│   └── (test files)
├── gui_main.py                       # GUI entry point
├── main.py                           # CLI entry point
├── requirements.txt                  # Dependencies
├── README.md                         # This file
└── LICENSE                           # MIT License
```

---

## 🧪 Testing

### Run Test Suite

```bash
# Install test dependencies
pip install pytest pytest-cov

# Run all tests
pytest tests/

# Run with coverage
pytest --cov=src tests/
```

### Manual Testing

See [TEST_PLAN.md](TEST_PLAN.md) for comprehensive test cases.

---


## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

## 🙏 Acknowledgments

- Built with [PySide6](https://doc.qt.io/qtforpython/) for the GUI
- Excel processing via [pandas](https://pandas.pydata.org/) and [openpyxl](https://openpyxl.readthedocs.io/)
- Inspired by real-world insurance and finance reconciliation needs

---

## 📧 Support

- **Issues**: [GitHub Issues](https://github.com/yourusername/excel-comparison-tool/issues)
- **Discussions**: [GitHub Discussions](https://github.com/yourusername/excel-comparison-tool/discussions)
- **Email**: pvela23@gmail.com

---

## 📊 Stats

![GitHub stars](https://img.shields.io/github/stars/yourusername/excel-comparison-tool?style=social)
![GitHub forks](https://img.shields.io/github/forks/yourusername/excel-comparison-tool?style=social)
![GitHub issues](https://img.shields.io/github/issues/yourusername/excel-comparison-tool)

---

**Made with ❤️ for data professionals who deserve better tools**

---

## Quick Links

- [Installation](#installation)
- [Usage Guide](#usage-guide)
- [Features](#features-in-detail)
- [Troubleshooting](#troubleshooting)
- [Contributing](#contributing)
- [License](#license)
