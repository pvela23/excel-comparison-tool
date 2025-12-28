# Quick Start Guide

Get up and running with Excel Comparison Tool in 5 minutes.

## 📦 Installation (One-Time Setup)

### Step 1: Install Python
- Download Python 3.8+ from [python.org](https://www.python.org/downloads/)
- During installation, check "Add Python to PATH"

### Step 2: Download the Tool
```bash
# Option A: Download ZIP from GitHub
# Extract to: C:\Tools\excel-comparison-tool\

# Option B: Git clone
git clone https://github.com/yourusername/excel-comparison-tool.git
cd excel-comparison-tool
```

### Step 3: Install Dependencies
```bash
# Windows Command Prompt
cd C:\Tools\excel-comparison-tool
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
```

**That's it!** Installation complete.

---

## 🚀 First Comparison (3 Minutes)

### Launch the Tool
```bash
# Make sure you're in the tool directory
cd C:\Tools\excel-comparison-tool
venv\Scripts\activate
python gui_main.py
```

A window opens:

```
┌─────────────────────────────────────┐
│     Excel Comparison Tool v1.0      │
└─────────────────────────────────────┘
```

### Compare Two Files

**Example: Insurance Policy Comparison**

1. **Select Files**
   - Click "Browse..." next to File A
   - Choose: `policies_january.xlsx`
   - Click "Browse..." next to File B
   - Choose: `policies_february.xlsx`

2. **Select Keys**
   - Check: ☑ `Policy Number`
   - Check: ☑ `Effective Date`
   - (These columns identify unique rows)

3. **Compare**
   - Click the big blue "Compare Files" button
   - Wait 2-5 seconds
   - Report opens automatically!

### Read the Report

Excel file opens with 3 sheets:

1. **Summary** - Quick overview
   ```
   Keys in common: 1,248
   Modified rows: 23
   Added rows: 15
   ```

2. **Aligned Diff** - Detailed comparison
   - Green rows = New policies
   - Red rows = Cancelled policies
   - Yellow cells = Changes
   - White rows = No change

3. **Legend** - Color guide & settings used

**Done!** You just compared two Excel files.

---

## 🎯 Common Scenarios

### Scenario 1: Simple Report Comparison
**Goal**: Compare last month's report to this month's

```
Files:
- sales_report_november.xlsx
- sales_report_december.xlsx

Keys: Customer ID

Result: Shows which customers had sales changes
```

### Scenario 2: Multi-Row Data
**Goal**: Compare policy data where one policy has multiple coverages

```
Files:
- policy_data_old.xlsx
- policy_data_new.xlsx

Keys: Policy Number + Coverage Type

Result: Shows which coverages were added/removed per policy
```

### Scenario 3: Invoice Reconciliation
**Goal**: Verify invoice system export vs. manual adjustments

```
Files:
- invoices_system_export.xlsx
- invoices_with_corrections.xlsx

Keys: Invoice Number + Line Item

Alignment: Secondary Sort by Line Item

Result: Shows exactly which line items changed
```

---

## 💡 Tips & Tricks

### Tip 1: Use Drag & Drop
Instead of clicking Browse:
- Drag 2 Excel files onto the window
- Files load automatically
- Faster workflow!

### Tip 2: Filter Large Column Lists
If you have 50+ columns:
- Type in the "Filter columns..." box
- Only matching columns show
- Example: Type "date" to see all date columns

### Tip 3: Keyboard Shortcut
After selecting keys, press `Ctrl+Enter` to compare immediately

### Tip 4: Save Time with Select All
If you need many key columns:
- Click "Select All" button
- Uncheck the few you don't need
- Faster than checking 20 boxes individually

---

## ❓ FAQ

**Q: Can I compare files with different column names?**
A: Currently no - column names must match. Column mapping coming in v1.1.

**Q: What's the maximum file size?**
A: Tested up to 1M rows. Files over 500K show a warning but work.

**Q: Can I compare more than 2 files?**
A: Not yet - pairwise comparison only. Compare A vs B, then B vs C.

**Q: Does it work on Mac/Linux?**
A: Yes! GUI works on all platforms. Auto-open works best on Windows.

**Q: Where are reports saved?**
A: Same folder as the tool. Named: `comparison_report_YYYYMMDD_HHMMSS.xlsx`

**Q: Can I automate this?**
A: Yes! Use `main.py` for CLI mode. See README for examples.

---

## 🆘 Troubleshooting

### "No module named 'pandas'"
```bash
# Make sure virtual environment is activated
venv\Scripts\activate

# Reinstall dependencies
pip install -r requirements.txt
```

### "Cannot access file (may be open in Excel)"
- Close Excel before loading the file
- Or save a copy and load the copy

### "No common columns"
- Check that column names match exactly (case-sensitive)
- Column "Policy" ≠ "policy"
- Remove extra spaces from column names

### Tool won't start
```bash
# Check Python version
python --version
# Should be 3.10 or higher

# Try using 'py' instead
py gui_main.py
```

---

## 📚 Next Steps

✅ You've completed your first comparison!

**Learn More:**
- Read the [full README](README.md) for all features
- Check [examples](examples/) folder for sample files
- See [advanced usage](README.md#advanced-usage) for automation

**Get Help:**
- Open an [issue](https://github.com/yourusername/excel-comparison-tool/issues)
- Check [discussions](https://github.com/yourusername/excel-comparison-tool/discussions)

---

**Happy comparing! 🎉**

Need help? Email: your.email@example.com
