# Excel Comparison Tool - Test Suite Guide

## Overview

This test suite provides comprehensive coverage for the Excel Comparison Tool project, including unit tests, integration tests, and edge case handling.

## Test Structure

### Test Files

1. **test_comparison_engine.py** - Tests for the core comparison engine
   - `ComparisonConfig` creation and validation
   - Basic comparison scenarios (identical, added, removed, modified rows)
   - Composite key comparisons
   - Data normalization (whitespace trimming, case sensitivity)
   - Different alignment methods
   - Edge cases (NaN values, empty DataFrames, unicode characters)
   - Complex real-world scenarios

2. **test_report_generator.py** - Tests for Excel report generation
   - Report creation and file saving
   - Sheet structure and content validation
   - Color coding and formatting
   - Various row status types
   - Unicode and large dataset handling

3. **test_integration.py** - End-to-end integration tests
   - Complete workflow: load → compare → report
   - Composite key workflows
   - Multi-row per key scenarios
   - Data normalization integration
   - Different alignment methods
   - Real-world business scenarios (insurance, financial transactions)

4. **conftest.py** - Pytest fixtures and configuration
   - Sample DataFrames for testing
   - Temporary file handling
   - Random state initialization

5. **test_utils.py** - Testing utilities and helpers
   - Test data generators
   - Result validators
   - File generation helpers

## Running Tests

### Prerequisites

```bash
pip install pytest pandas openpyxl numpy
```

### Run All Tests

```bash
pytest
```

### Run Specific Test File

```bash
pytest tests/test_comparison_engine.py
pytest tests/test_report_generator.py
pytest tests/test_integration.py
```

### Run Specific Test Class

```bash
pytest tests/test_comparison_engine.py::TestComparisonEngineBasic
```

### Run Specific Test

```bash
pytest tests/test_comparison_engine.py::TestComparisonEngineBasic::test_identical_dataframes
```

### Run with Verbose Output

```bash
pytest -v
```

### Run with Coverage Report

```bash
pip install pytest-cov
pytest --cov=src --cov-report=html
```

### Run with Specific Markers

```bash
pytest -m unit
pytest -m integration
pytest -m "not slow"
```

## Test Coverage

### Comparison Engine (~80 tests)
- ✅ Configuration creation and validation
- ✅ Basic comparison scenarios
- ✅ Row status detection (match, modified, added, removed, new key, removed key)
- ✅ Single and composite key comparisons
- ✅ Multi-row per key handling
- ✅ Data normalization (case sensitivity, whitespace trimming)
- ✅ Alignment methods (position, secondary sort)
- ✅ Edge cases (NaN, unicode, large datasets)
- ✅ Error handling (missing columns, invalid inputs)

### Report Generator (~35 tests)
- ✅ Report creation and saving
- ✅ Sheet structure validation
- ✅ Color coding
- ✅ Summary sheet content
- ✅ Aligned diff sheet
- ✅ Legend sheet
- ✅ Various row statuses
- ✅ Unicode and special characters
- ✅ Large dataset handling

### Integration (~20 tests)
- ✅ End-to-end workflows
- ✅ Composite key workflows
- ✅ Multi-row per key scenarios
- ✅ Data normalization integration
- ✅ Alignment method integration
- ✅ Real-world business scenarios
- ✅ Report generation integration
- ✅ Error handling

## Test Examples

### Example: Testing Basic Comparison

```python
def test_identical_dataframes():
    df_a = pd.DataFrame({
        'ID': [1, 2, 3],
        'Name': ['Alice', 'Bob', 'Charlie'],
        'Value': [100, 200, 300]
    })
    df_b = df_a.copy()
    
    config = ComparisonConfig(key_columns=['ID'])
    engine = ComparisonEngine(config)
    result = engine.compare(df_a, df_b)
    
    assert result.summary['match_count'] == 3
    assert result.summary['modified_count'] == 0
```

### Example: Testing Modified Rows

```python
def test_modified_rows():
    df_a = pd.DataFrame({'ID': [1, 2], 'Value': [100, 200]})
    df_b = pd.DataFrame({'ID': [1, 2], 'Value': [100, 250]})
    
    config = ComparisonConfig(key_columns=['ID'])
    engine = ComparisonEngine(config)
    result = engine.compare(df_a, df_b)
    
    assert result.summary['match_count'] == 1
    assert result.summary['modified_count'] == 1
```

### Example: Using Test Fixtures

```python
def test_with_fixture(sample_dataframe_a, sample_dataframe_b):
    config = ComparisonConfig(key_columns=['ID'])
    engine = ComparisonEngine(config)
    result = engine.compare(sample_dataframe_a, sample_dataframe_b)
    
    # Test assertions here
    assert result.summary['keys_in_common'] > 0
```

## Common Test Patterns

### Pattern: Testing Edge Cases
```python
def test_nan_values():
    df_a = pd.DataFrame({'ID': [1], 'Value': [np.nan]})
    df_b = pd.DataFrame({'ID': [1], 'Value': [np.nan]})
    # NaN should equal NaN in this context
```

### Pattern: Testing Error Conditions
```python
def test_missing_key_column():
    config = ComparisonConfig(key_columns=['NonExistent'])
    engine = ComparisonEngine(config)
    
    with pytest.raises(KeyError):
        engine.compare(df_a, df_b)
```

### Pattern: Testing Real-World Scenarios
```python
def test_insurance_policy_comparison():
    # Create realistic data with multiple coverages per policy
    # Test comparison results
```

## Continuous Integration

Tests are designed to run in CI/CD pipelines. All tests:
- Are isolated and don't depend on external resources
- Use temporary files for Excel generation
- Clean up resources automatically
- Run quickly (majority complete in < 100ms)

## Troubleshooting

### Import Errors
Ensure the project root is in PYTHONPATH:
```bash
export PYTHONPATH="${PYTHONPATH}:$(pwd)"
```

### Temporary File Issues
Tests use `tempfile` module for automatic cleanup. If cleanup fails, check:
- Disk space availability
- File permissions in temp directory
- No lingering Excel processes

### DataFrame Comparison Issues
Remember:
- NaN != NaN by default (tests handle this)
- Case sensitivity is configurable
- Whitespace trimming is configurable
- Numeric types (int vs float) should match or be configured

## Adding New Tests

1. Create test class in appropriate file
2. Follow naming convention: `test_<description>`
3. Use descriptive docstrings
4. Add assertions with clear messages
5. Clean up any created resources
6. Consider edge cases

Example:
```python
class TestNewFeature:
    """Test description"""
    
    def test_feature_behavior(self):
        """Test specific behavior"""
        # Arrange
        data = prepare_test_data()
        
        # Act
        result = perform_action(data)
        
        # Assert
        assert expected_condition(result)
```

## Performance Notes

- Most tests complete in < 100ms
- Large dataset tests (500K+ rows) may take longer
- Report generation tests involve file I/O
- Use `-v` flag for detailed timing information

## Test Dependencies

- **pandas**: DataFrame creation and manipulation
- **numpy**: Numeric operations and NaN handling
- **openpyxl**: Excel file validation
- **pytest**: Test framework and fixtures
- **tempfile**: Temporary file handling
