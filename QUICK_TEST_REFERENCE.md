# Quick Test Reference

## Run Tests Quickly

```bash
# All tests
pytest

# With coverage
pytest --cov=src

# Specific test file
pytest tests/test_comparison_engine.py

# Specific test
pytest tests/test_comparison_engine.py::TestComparisonEngineBasic::test_identical_dataframes

# With HTML report
pytest --cov=src --cov-report=html
```

## Windows Batch Commands

```batch
test.bat all           # All tests
test.bat fast          # Fast (no coverage)
test.bat coverage      # With coverage
test.bat html          # HTML coverage report
test.bat engine        # Engine tests only
test.bat report        # Report tests only
test.bat integration   # Integration tests only
test.bat verbose       # Verbose output
```

## What's Tested

### Comparison Engine (27 tests)
- ✅ Configuration and setup
- ✅ Single & composite key comparisons
- ✅ Row status detection
- ✅ Data normalization
- ✅ Alignment methods
- ✅ Edge cases & errors

### Report Generator (14 tests)
- ✅ Excel file creation
- ✅ Sheet structure
- ✅ Color coding
- ✅ Data formatting

### Integration (14 tests)
- ✅ End-to-end workflows
- ✅ Real-world scenarios
- ✅ Multi-component testing

## Test Coverage

```
Total: 55 tests ✅
Coverage: 97%
Time: ~1 second
```

## Test Files

- `tests/test_comparison_engine.py` - Core engine tests
- `tests/test_report_generator.py` - Report generation tests
- `tests/test_integration.py` - End-to-end tests
- `tests/conftest.py` - Test configuration & fixtures
- `tests/test_utils.py` - Test helpers & utilities

## Documentation

- `tests/TEST_GUIDE.md` - Complete testing guide
- `TEST_SUMMARY.md` - Test results & metrics
- `TESTING_COMPLETE.md` - Implementation details

## Key Features Tested

- Key-based comparisons (single & composite)
- Multi-row per key handling
- All row status types
- Case sensitivity & whitespace handling
- Position & secondary sort alignment
- NaN handling
- Large dataset support
- Unicode character support
- Error conditions
- Real-world business scenarios

## Adding New Tests

1. Create test in appropriate file
2. Follow naming: `test_<description>`
3. Use AAA pattern (Arrange-Act-Assert)
4. Add docstring
5. Use fixtures for common data
6. Clean up resources

Example:
```python
def test_feature():
    """Test description"""
    # Arrange
    df = pd.DataFrame({'ID': [1], 'Value': [100]})
    
    # Act
    result = perform_comparison(df)
    
    # Assert
    assert result.summary['match_count'] == 1
```

## Troubleshooting

**Import Error?**
```bash
set PYTHONPATH=%CD%
```

**Need to install pytest?**
```bash
pip install pytest pytest-cov
```

**Want verbose output?**
```bash
pytest -v
```

**Only run fast tests?**
```bash
pytest -m "not slow"
```

## Quick Stats

| Metric | Value |
|--------|-------|
| Total Tests | 55 |
| Passing | 55 ✅ |
| Code Coverage | 97% |
| Execution Time | ~1s |
| Test Files | 5 |
| Avg Test Time | ~18ms |

---

For more details, see [tests/TEST_GUIDE.md](tests/TEST_GUIDE.md) or [TEST_SUMMARY.md](TEST_SUMMARY.md)
