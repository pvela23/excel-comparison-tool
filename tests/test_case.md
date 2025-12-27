TEST CASES
🟢 CATEGORY 1: Basic Functionality
Test 1.1: Simple Comparison - Match
Steps:
Launch tool
Browse File A: test_simple_a.xlsx
Browse File B: test_simple_a.xlsx (same file)
Select key: Policy
Click Compare
Expected:
✅ All 3 rows show as MATCH
✅ 0 modified, 0 added, 0 removed
✅ Report opens
✅ Green status message

Test 1.2: Simple Comparison - With Changes
Steps:
Load test_simple_a.xlsx and test_simple_b.xlsx
Select key: Policy
Click Compare
Expected:
✅ 1 MATCH (12346)
✅ 1 MODIFIED (12345 - Premium changed)
✅ 1 REMOVED_KEY (12347)
✅ 1 NEW_KEY (12348)
✅ Summary shows correct counts
✅ Excel opens with color-coded rows

Test 1.3: Multi-Row Per Key
Steps:
Load test_multirow_a.xlsx and test_multirow_b.xlsx
Select keys: Policy + EFF
Click Compare
Expected:
✅ Policy 12345 shows:
1 MODIFIED (Auto: 500→550)
1 MATCH (Home: 800)
1 ADDED_ROW (Life: 200)
✅ Policy 12346 shows: 1 MATCH
✅ Rows grouped by policy in report

🟡 CATEGORY 2: Alignment Methods
Test 2.1: Position-Based Alignment
Steps:
Load test_multirow_a.xlsx and test_multirow_b.xlsx
Select key: Policy
Alignment: Position-based
Compare
Expected:
✅ For Policy 12345:
Row 1A (Auto) → Row 1B (Auto) = MODIFIED
Row 2A (Home) → Row 2B (Home) = MATCH
Row 3B (Life) = ADDED

Test 2.2: Secondary Sort Alignment
Steps:
Load test_multirow_a.xlsx and test_multirow_b.xlsx
Select key: Policy
Alignment: Secondary Sort Column
Sort by: Coverage
Compare
Expected:
✅ Dropdown shows "Sort By Column" field
✅ Comparison sorts by Coverage before matching
✅ Auto→Auto, Home→Home, Life→(added)

Test 2.3: Secondary Sort - No Column Selected
Steps:
Select Secondary Sort alignment
Don't select a sort column
Click Compare
Expected:
✅ Warning: "Please select a column to sort by"
✅ Comparison doesn't start

🔵 CATEGORY 3: File Loading
Test 3.1: Multi-Sheet Selection
Steps:
Browse test_multisheet_a.xlsx
Dialog appears asking to select sheet
Expected:
✅ Dialog shows: "Select Sheet" with Sheet1, Sheet2
✅ Can select Sheet1 or Sheet2
✅ File display shows: path [Sheet1]
✅ Tooltip shows sheet name

Test 3.2: Large File Warning
Steps:
Browse test_large.xlsx (600k rows)
Expected:
✅ Warning dialog: "This file has 600,000 rows"
✅ Option to continue or cancel
✅ If cancel → file not loaded
✅ If continue → loads normally

Test 3.3: Empty File
Steps:
Browse test_empty.xlsx
Expected:
✅ Warning: "The selected sheet appears to be empty"
✅ File not loaded

Test 3.4: File Open in Excel
Steps:
Open test_simple_a.xlsx in Excel
Try to browse it in the tool
Expected:
✅ Error: "Cannot access file (it may be open in Excel)"
✅ Message: "Please close the file and try again"

Test 3.5: Invalid File Format
Steps:
Try to browse a .txt or .pdf file
Expected:
✅ File picker only shows .xlsx, .xls, .xlsm
✅ Other files not selectable

Test 3.6: Non-Existent File
Steps:
Manually type invalid path in code/test
Try to load
Expected:
✅ Error: "Could not find the file"

🟠 CATEGORY 4: Key Selection
Test 4.1: No Keys Selected
Steps:
Load two files
Don't check any key columns
Click Compare
Expected:
✅ Warning: "Please select at least one key column"
✅ Comparison doesn't start

Test 4.2: No Common Columns
Steps:
Load test_no_common.xlsx (ColA, ColB)
Load test_no_common2.xlsx (ColX, ColY)
Expected:
✅ Warning: "These files have no columns in common"
✅ Shows first 5 columns from each file
✅ Compare button stays disabled

Test 4.3: Filter Columns
Steps:
Load files with many columns (10+)
Type "pol" in filter box
Expected:
✅ Only "Policy" checkbox visible
✅ Label shows: "Showing 1 of 10 columns"
✅ Clear filter → all columns reappear

Test 4.4: Select All / Deselect All
Steps:
Load files
Click "Select All"
Click "Deselect All"
Expected:
✅ Select All → All checkboxes checked
✅ Count shows: "Selected: 10"
✅ Deselect All → All unchecked
✅ Count shows: "Selected: 0"

Test 4.5: Select All with Filter
Steps:
Type filter text (shows 3 of 10)
Click "Select All"
Expected:
✅ Only visible (filtered) columns selected
✅ Hidden columns remain unchecked

🟣 CATEGORY 5: Options & Settings
Test 5.1: Case Sensitive ON
Setup:
File A: Policy = "ABC"
File B: Policy = "abc"
Steps:
Check "Case Sensitive"
Compare
Expected:
✅ Rows shown as MODIFIED (ABC ≠ abc)

Test 5.2: Case Sensitive OFF
Same setup as 5.1 Steps:
Uncheck "Case Sensitive"
Compare
Expected:
✅ Rows shown as MATCH (abc = ABC when ignoring case)

Test 5.3: Trim Whitespace ON
Setup:
File A: Policy = "12345 "  (trailing space)
File B: Policy = "12345"
Steps:
Trim Whitespace = ON (default)
Compare
Expected:
✅ Rows shown as MATCH

Test 5.4: Trim Whitespace OFF
Same setup as 5.3 Steps:
Uncheck "Trim Whitespace"
Compare
Expected:
✅ Rows shown as MODIFIED ("12345 " ≠ "12345")

Test 5.5: Settings Persistence
Steps:
Check "Case Sensitive"
Close tool
Reopen tool
Expected:
✅ "Case Sensitive" still checked
✅ Window size/position restored

🔴 CATEGORY 6: Drag & Drop
Test 6.1: Drop Single File
Steps:
Drag test_simple_a.xlsx onto window
Expected:
✅ Loads into File A
✅ Status bar shows file loaded

Test 6.2: Drop Two Files
Steps:
Select two Excel files
Drag both onto window
Expected:
✅ First file → File A
✅ Second file → File B
✅ Dialog: "Files Loaded: File A: ..., File B: ..."

Test 6.3: Drop Non-Excel File
Steps:
Drag a .txt file onto window
Expected:
✅ Warning: "Please drop Excel files (.xlsx, .xls, .xlsm)"

Test 6.4: Drop When File A Already Loaded
Steps:
Load File A via Browse
Drag another file onto window
Expected:
✅ New file goes into File B
✅ File A unchanged

⚫ CATEGORY 7: Keyboard & Shortcuts
Test 7.1: Ctrl+Enter to Compare
Steps:
Load files, select keys
Press Ctrl+Enter
Expected:
✅ Comparison starts (same as clicking button)

Test 7.2: Enter in Filter Box
Steps:
Type in filter box
Press Enter
Expected:
✅ Filter applies
✅ Doesn't trigger comparison

⚪ CATEGORY 8: Results & Reporting
Test 8.1: Results Dialog - Show Details
Steps:
Complete comparison
Click "Show Details" in dialog
Expected:
✅ Expanded section shows:
Full statistics
Configuration used
Source file paths with sheets
Report location

Test 8.2: Open Report Button
Steps:
Complete comparison
Click "Open Report" in dialog
Expected:
✅ Excel opens automatically (Windows)
✅ Report file displayed

Test 8.3: Report File Location
Steps:
Complete comparison
Check file system
Expected:
✅ File exists in tool directory
✅ Named: comparison_report_YYYYMMDD_HHMMSS.xlsx
✅ Has 3 sheets: Summary, Aligned Diff, Legend

Test 8.4: Multiple Comparisons (No Close)
Steps:
Compare files A1 & B1
Without closing, change to A2 & B2
Compare again
Expected:
✅ Second comparison works
✅ No sheet name errors
✅ New report file created with new timestamp

🟤 CATEGORY 9: Edge Cases
Test 9.1: Identical Files Different Sheets
Steps:
Load test_multisheet_a.xlsx Sheet1
Load same file Sheet2
Expected:
✅ Compares different sheets from same file
✅ Works normally

Test 9.2: 100% Match (No Differences)
Steps:
Load same file twice, same sheet
Expected:
✅ All rows = MATCH
✅ 0 modified, 0 added, 0 removed
✅ Report still generated

Test 9.3: 100% Different (No Common Keys)
Steps:
File A: Policy 001, 002, 003
File B: Policy 004, 005, 006
Expected:
✅ All File A keys = REMOVED_KEY
✅ All File B keys = NEW_KEY
✅ 0 keys in common

Test 9.4: Composite Key (3+ columns)
Steps:
Select 3 keys: Policy + EFF + Coverage
Compare
Expected:
✅ Comparison uses all 3 columns as composite key
✅ Report shows all 3 in key columns

Test 9.5: Column Names with Special Characters
Setup:
Columns: "Policy #", "Eff. Date", "Premium ($)"
Expected:
✅ Loads without error
✅ Checkboxes display correctly
✅ Comparison works

Test 9.6: Very Long Column Names
Setup:
Column: "This_Is_A_Very_Long_Column_Name_That_Goes_On_And_On"
Expected:
✅ Checkbox shows full name
✅ UI doesn't break
✅ Comparison works

🔶 CATEGORY 10: Performance & Progress
Test 10.1: Progress Bar Visibility
Steps:
Start comparison
Watch progress bar
Expected:
✅ Progress bar appears immediately
✅ Shows indeterminate animation (spinning)
✅ Status bar updates: "Comparing..." → "Generating report..."

Test 10.2: Button States During Comparison
Steps:
Start comparison
Check button states
Expected:
✅ Compare button = disabled
✅ Config section = disabled
✅ Can't start second comparison

Test 10.3: Comparison Time Display
Steps:
Compare small files (< 1 sec)
Check results dialog
Expected:
✅ Shows: "0.05 seconds" (or similar)

Test 10.4: Long Comparison Time
Steps:
Compare large files (> 60 seconds)
Expected:
✅ Shows: "2 min 15.3 sec" format

🔷 CATEGORY 11: Error Recovery
Test 11.1: Crash During Comparison
Steps:
Simulate error (modify code to raise exception mid-comparison)
Expected:
✅ Error dialog appears
✅ UI returns to usable state
✅ Can try comparison again

Test 11.2: Disk Full (Can't Write Report)
Setup: Fill disk to capacity Steps:
Try to run comparison
Expected:
✅ Error message about disk space
✅ Tool doesn't crash

Test 11.3: Excel File Corrupted
Steps:
Create corrupted .xlsx (corrupt zip file)
Try to load
Expected:
✅ Error: "Invalid Excel file format"
✅ Doesn't crash

📊 TEST RESULTS TEMPLATE
Use this to track results:
TEST ID | Test Name                  | Status | Notes
--------|----------------------------|--------|-------
1.1     | Simple Match               | ✅     | 
1.2     | Simple With Changes        | ✅     |
1.3     | Multi-Row Per Key          | ✅     |
2.1     | Position-Based             | ❌     | Bug: ...
2.2     | Secondary Sort             | ⏸️     | Skipped
...

🎯 PRIORITY ORDER
Critical (Must Pass):
1.1, 1.2, 1.3 (Basic functionality)
2.1, 2.2 (Alignment methods)
4.1, 4.2 (Key validation)
8.4 (Multiple comparisons)
High Priority:
3.x (All file loading)
4.x (All key selection)
8.x (Results reporting)
Medium Priority:
5.x (Options)
6.x (Drag & drop)
9.x (Edge cases)
Low Priority:
7.x (Shortcuts - nice to have)
10.x (Progress feedback)

🐛 BUG REPORTING TEMPLATE
When you find a bug:
BUG #: 001
Test Case: 2.2 (Secondary Sort)
Steps to Reproduce:
1. ...
2. ...

Expected: Sorts by Coverage column
Actual: Error message appears
Error: "KeyError: Coverage"

Severity: High
Screenshots: bug001.png

Run through these tests and let me know:
Which tests pass ✅
Which tests fail ❌
Any unexpected behavior
Any crashes or errors

