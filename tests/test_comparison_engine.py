"""
Unit tests for comparison_engine module
Tests core comparison logic with various scenarios
"""

import pytest
import pandas as pd
import numpy as np
from src.core import (
    ComparisonEngine,
    ComparisonConfig,
    RowStatus,
    AlignmentMethod
)


class TestComparisonConfig:
    """Test ComparisonConfig dataclass"""
    
    def test_config_creation_with_defaults(self):
        """Test creating config with default values"""
        config = ComparisonConfig(key_columns=['ID'])
        assert config.key_columns == ['ID']
        assert config.alignment_method == AlignmentMethod.POSITION
        assert config.case_sensitive == False
        assert config.trim_whitespace == True
    
    def test_config_creation_with_custom_values(self):
        """Test creating config with custom values"""
        config = ComparisonConfig(
            key_columns=['ID', 'Name'],
            alignment_method=AlignmentMethod.SECONDARY_SORT,
            secondary_sort_column='Date',
            case_sensitive=True,
            trim_whitespace=False
        )
        assert config.key_columns == ['ID', 'Name']
        assert config.alignment_method == AlignmentMethod.SECONDARY_SORT
        assert config.secondary_sort_column == 'Date'
        assert config.case_sensitive == True
        assert config.trim_whitespace == False


class TestComparisonEngineBasic:
    """Test basic comparison engine functionality"""
    
    def test_engine_creation(self):
        """Test creating comparison engine"""
        config = ComparisonConfig(key_columns=['ID'])
        engine = ComparisonEngine(config)
        assert engine.config == config
    
    def test_identical_dataframes(self):
        """Test comparing identical DataFrames"""
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
        assert result.summary['added_row_count'] == 0
        assert result.summary['removed_row_count'] == 0
        assert result.summary['keys_in_common'] == 3
        assert result.summary['keys_only_in_a'] == 0
        assert result.summary['keys_only_in_b'] == 0
    
    def test_missing_key_column(self):
        """Test error when key column is missing"""
        df_a = pd.DataFrame({'ID': [1, 2], 'Name': ['Alice', 'Bob']})
        df_b = pd.DataFrame({'ID': [1, 2], 'Value': [100, 200]})
        
        config = ComparisonConfig(key_columns=['NonExistent'])
        engine = ComparisonEngine(config)
        
        with pytest.raises(KeyError):
            engine.compare(df_a, df_b)


class TestComparisonEngineRows:
    """Test row comparison scenarios"""
    
    def test_added_rows(self):
        """Test detecting added rows"""
        df_a = pd.DataFrame({
            'ID': [1, 2],
            'Name': ['Alice', 'Bob'],
            'Value': [100, 200]
        })
        df_b = pd.DataFrame({
            'ID': [1, 2, 3],
            'Name': ['Alice', 'Bob', 'Charlie'],
            'Value': [100, 200, 300]
        })
        
        config = ComparisonConfig(key_columns=['ID'])
        engine = ComparisonEngine(config)
        result = engine.compare(df_a, df_b)
        
        # New key (ID=3) is counted in new_key_count, not added_row_count
        assert result.summary['new_key_count'] == 1
        assert result.summary['removed_row_count'] == 0
        assert result.summary['match_count'] == 2
        assert result.summary['keys_only_in_b'] == 1
    
    def test_removed_rows(self):
        """Test detecting removed rows"""
        df_a = pd.DataFrame({
            'ID': [1, 2, 3],
            'Name': ['Alice', 'Bob', 'Charlie'],
            'Value': [100, 200, 300]
        })
        df_b = pd.DataFrame({
            'ID': [1, 2],
            'Name': ['Alice', 'Bob'],
            'Value': [100, 200]
        })
        
        config = ComparisonConfig(key_columns=['ID'])
        engine = ComparisonEngine(config)
        result = engine.compare(df_a, df_b)
        
        # Removed key (ID=3) is counted in removed_key_count, not removed_row_count
        assert result.summary['removed_key_count'] == 1
        assert result.summary['added_row_count'] == 0
        assert result.summary['match_count'] == 2
        assert result.summary['keys_only_in_a'] == 1
    
    def test_modified_rows(self):
        """Test detecting modified rows"""
        df_a = pd.DataFrame({
            'ID': [1, 2],
            'Name': ['Alice', 'Bob'],
            'Value': [100, 200]
        })
        df_b = pd.DataFrame({
            'ID': [1, 2],
            'Name': ['Alice', 'Bobby'],  # Modified
            'Value': [100, 250]  # Modified
        })
        
        config = ComparisonConfig(key_columns=['ID'])
        engine = ComparisonEngine(config)
        result = engine.compare(df_a, df_b)
        
        assert result.summary['match_count'] == 1
        assert result.summary['modified_count'] == 1
    
    def test_multi_row_per_key(self):
        """Test multiple rows with same key"""
        df_a = pd.DataFrame({
            'Policy': ['P001', 'P001', 'P002'],
            'Coverage': ['A', 'B', 'A'],
            'Premium': [100, 50, 200]
        })
        df_b = pd.DataFrame({
            'Policy': ['P001', 'P001', 'P002', 'P002'],
            'Coverage': ['A', 'B', 'A', 'B'],
            'Premium': [100, 50, 200, 150]
        })
        
        config = ComparisonConfig(key_columns=['Policy'])
        engine = ComparisonEngine(config)
        result = engine.compare(df_a, df_b)
        
        assert result.summary['keys_in_common'] == 2
        assert result.summary['added_row_count'] == 1
        assert result.summary['match_count'] == 3


class TestComparisonEngineCompositeKeys:
    """Test composite key scenarios"""
    
    def test_composite_key_comparison(self):
        """Test comparison with composite key"""
        df_a = pd.DataFrame({
            'PolicyID': [1, 1, 2],
            'CoverageID': ['A', 'B', 'A'],
            'Premium': [100, 50, 200]
        })
        df_b = pd.DataFrame({
            'PolicyID': [1, 1, 2],
            'CoverageID': ['A', 'B', 'A'],
            'Premium': [100, 60, 200]  # Middle row modified
        })
        
        config = ComparisonConfig(key_columns=['PolicyID', 'CoverageID'])
        engine = ComparisonEngine(config)
        result = engine.compare(df_a, df_b)
        
        assert result.summary['keys_in_common'] == 3
        assert result.summary['match_count'] == 2
        assert result.summary['modified_count'] == 1
    
    def test_composite_key_with_new_keys(self):
        """Test new composite keys in file B"""
        df_a = pd.DataFrame({
            'PolicyID': [1, 2],
            'CoverageID': ['A', 'A'],
            'Premium': [100, 200]
        })
        df_b = pd.DataFrame({
            'PolicyID': [1, 2, 3],
            'CoverageID': ['A', 'A', 'A'],
            'Premium': [100, 200, 300]
        })
        
        config = ComparisonConfig(key_columns=['PolicyID', 'CoverageID'])
        engine = ComparisonEngine(config)
        result = engine.compare(df_a, df_b)
        
        assert result.summary['keys_only_in_b'] == 1
        assert result.summary['new_key_count'] == 1


class TestDataNormalization:
    """Test data normalization features"""
    
    def test_trim_whitespace(self):
        """Test trimming whitespace"""
        df_a = pd.DataFrame({
            'ID': [1],
            'Name': [' Alice ']
        })
        df_b = pd.DataFrame({
            'ID': [1],
            'Name': ['Alice']
        })
        
        config = ComparisonConfig(
            key_columns=['ID'],
            trim_whitespace=True
        )
        engine = ComparisonEngine(config)
        result = engine.compare(df_a, df_b)
        
        assert result.summary['match_count'] == 1
        assert result.summary['modified_count'] == 0
    
    def test_case_sensitivity(self):
        """Test case sensitivity option"""
        df_a = pd.DataFrame({
            'ID': [1],
            'Name': ['Alice']
        })
        df_b = pd.DataFrame({
            'ID': [1],
            'Name': ['ALICE']
        })
        
        # Case insensitive
        config = ComparisonConfig(
            key_columns=['ID'],
            case_sensitive=False
        )
        engine = ComparisonEngine(config)
        result = engine.compare(df_a, df_b)
        assert result.summary['match_count'] == 1
        
        # Case sensitive
        config = ComparisonConfig(
            key_columns=['ID'],
            case_sensitive=True
        )
        engine = ComparisonEngine(config)
        result = engine.compare(df_a, df_b)
        assert result.summary['modified_count'] == 1
    
    def test_whitespace_disabled(self):
        """Test with whitespace trimming disabled"""
        df_a = pd.DataFrame({
            'ID': [1],
            'Name': [' Alice ']
        })
        df_b = pd.DataFrame({
            'ID': [1],
            'Name': ['Alice']
        })
        
        config = ComparisonConfig(
            key_columns=['ID'],
            trim_whitespace=False
        )
        engine = ComparisonEngine(config)
        result = engine.compare(df_a, df_b)
        
        assert result.summary['modified_count'] == 1


class TestAlignmentMethods:
    """Test different alignment methods"""
    
    def test_position_based_alignment(self):
        """Test position-based alignment (default)"""
        df_a = pd.DataFrame({
            'Policy': ['P001', 'P001', 'P001'],
            'Coverage': ['A', 'B', 'C'],
            'Premium': [100, 50, 25]
        })
        df_b = pd.DataFrame({
            'Policy': ['P001', 'P001', 'P001'],
            'Coverage': ['X', 'Y', 'Z'],
            'Premium': [110, 55, 30]
        })
        
        config = ComparisonConfig(
            key_columns=['Policy'],
            alignment_method=AlignmentMethod.POSITION
        )
        engine = ComparisonEngine(config)
        result = engine.compare(df_a, df_b)
        
        assert result.summary['keys_in_common'] == 1
        assert result.summary['modified_count'] == 3
    
    def test_secondary_sort_alignment(self):
        """Test secondary sort alignment"""
        df_a = pd.DataFrame({
            'Policy': ['P001', 'P001', 'P001'],
            'Date': ['2024-03-01', '2024-01-01', '2024-02-01'],
            'Premium': [100, 200, 150]
        })
        df_b = pd.DataFrame({
            'Policy': ['P001', 'P001', 'P001'],
            'Date': ['2024-01-01', '2024-02-01', '2024-03-01'],
            'Premium': [200, 150, 100]
        })
        
        config = ComparisonConfig(
            key_columns=['Policy'],
            alignment_method=AlignmentMethod.SECONDARY_SORT,
            secondary_sort_column='Date'
        )
        engine = ComparisonEngine(config)
        result = engine.compare(df_a, df_b)
        
        # Should match all three rows when sorted by date
        assert result.summary['match_count'] == 3


class TestEdgeCases:
    """Test edge cases and special scenarios"""
    
    def test_empty_dataframe_a(self):
        """Test with empty File A"""
        df_a = pd.DataFrame({
            'ID': pd.Series([], dtype=int),
            'Name': pd.Series([], dtype=str)
        })
        df_b = pd.DataFrame({
            'ID': [1, 2],
            'Name': ['Alice', 'Bob']
        })
        
        config = ComparisonConfig(key_columns=['ID'])
        engine = ComparisonEngine(config)
        result = engine.compare(df_a, df_b)
        
        # Keys only in B (new keys) instead of added_row_count
        assert result.summary['new_key_count'] == 2
        assert result.summary['keys_only_in_b'] == 2
    
    def test_empty_dataframe_b(self):
        """Test with empty File B"""
        df_a = pd.DataFrame({
            'ID': [1, 2],
            'Name': ['Alice', 'Bob']
        })
        df_b = pd.DataFrame({
            'ID': pd.Series([], dtype=int),
            'Name': pd.Series([], dtype=str)
        })
        
        config = ComparisonConfig(key_columns=['ID'])
        engine = ComparisonEngine(config)
        result = engine.compare(df_a, df_b)
        
        # Keys only in A (removed keys) instead of removed_row_count
        assert result.summary['removed_key_count'] == 2
        assert result.summary['keys_only_in_a'] == 2
    
    def test_nan_values(self):
        """Test handling of NaN values"""
        df_a = pd.DataFrame({
            'ID': [1, 2],
            'Value': [100.0, np.nan]
        })
        df_b = pd.DataFrame({
            'ID': [1, 2],
            'Value': [100.0, np.nan]
        })
        
        config = ComparisonConfig(key_columns=['ID'])
        engine = ComparisonEngine(config)
        result = engine.compare(df_a, df_b)
        
        # Both NaN should be treated as equal
        assert result.summary['match_count'] == 2
    
    def test_nan_vs_value(self):
        """Test NaN vs actual value"""
        df_a = pd.DataFrame({
            'ID': [1],
            'Value': [np.nan]
        })
        df_b = pd.DataFrame({
            'ID': [1],
            'Value': [100.0]
        })
        
        config = ComparisonConfig(key_columns=['ID'])
        engine = ComparisonEngine(config)
        result = engine.compare(df_a, df_b)
        
        assert result.summary['modified_count'] == 1
    
    def test_numeric_types(self):
        """Test comparison with different numeric types"""
        df_a = pd.DataFrame({
            'ID': [1],
            'Value': [100]  # Integer
        })
        df_b = pd.DataFrame({
            'ID': [1],
            'Value': [100.0]  # Float
        })
        
        config = ComparisonConfig(key_columns=['ID'])
        engine = ComparisonEngine(config)
        result = engine.compare(df_a, df_b)
        
        assert result.summary['match_count'] == 1
    
    def test_special_characters_in_keys(self):
        """Test keys with special characters"""
        df_a = pd.DataFrame({
            'ID': ['P-001-A', 'P-002-B'],
            'Value': [100, 200]
        })
        df_b = pd.DataFrame({
            'ID': ['P-001-A', 'P-002-B'],
            'Value': [100, 200]
        })
        
        config = ComparisonConfig(key_columns=['ID'])
        engine = ComparisonEngine(config)
        result = engine.compare(df_a, df_b)
        
        assert result.summary['match_count'] == 2
    
    def test_large_number_of_rows(self):
        """Test with larger dataset"""
        n = 1000
        df_a = pd.DataFrame({
            'ID': range(n),
            'Value': range(100, 100 + n)
        })
        df_b = df_a.copy()
        df_b.loc[500, 'Value'] = 999  # Modify one row
        
        config = ComparisonConfig(key_columns=['ID'])
        engine = ComparisonEngine(config)
        result = engine.compare(df_a, df_b)
        
        assert result.summary['match_count'] == n - 1
        assert result.summary['modified_count'] == 1
        assert result.summary['total_rows_compared'] == n
    
    def test_unicode_characters(self):
        """Test handling of unicode characters"""
        df_a = pd.DataFrame({
            'ID': [1, 2],
            'Name': ['José', 'François']
        })
        df_b = pd.DataFrame({
            'ID': [1, 2],
            'Name': ['José', 'François']
        })
        
        config = ComparisonConfig(key_columns=['ID'])
        engine = ComparisonEngine(config)
        result = engine.compare(df_a, df_b)
        
        assert result.summary['match_count'] == 2


class TestComparisonResults:
    """Test the structure of comparison results"""
    
    def test_result_contains_aligned_data(self):
        """Test that result contains aligned data"""
        df_a = pd.DataFrame({'ID': [1, 2], 'Value': [100, 200]})
        df_b = pd.DataFrame({'ID': [1, 2, 3], 'Value': [100, 200, 300]})
        
        config = ComparisonConfig(key_columns=['ID'])
        engine = ComparisonEngine(config)
        result = engine.compare(df_a, df_b)
        
        assert not result.aligned_data.empty
        assert 'status' in result.aligned_data.columns
    
    def test_result_metadata(self):
        """Test that result contains proper metadata"""
        df_a = pd.DataFrame({'ID': [1], 'Value': [100]})
        df_b = pd.DataFrame({'ID': [1], 'Value': [100]})
        
        config = ComparisonConfig(key_columns=['ID'])
        engine = ComparisonEngine(config)
        result = engine.compare(df_a, df_b)
        
        assert 'config' in result.comparison_metadata
        assert result.comparison_metadata['total_rows_a'] == 1
        assert result.comparison_metadata['total_rows_b'] == 1
    
    def test_keys_only_lists(self):
        """Test keys_only_in_a and keys_only_in_b lists"""
        df_a = pd.DataFrame({'ID': [1, 2, 3], 'Value': [100, 200, 300]})
        df_b = pd.DataFrame({'ID': [2, 3, 4], 'Value': [200, 300, 400]})
        
        config = ComparisonConfig(key_columns=['ID'])
        engine = ComparisonEngine(config)
        result = engine.compare(df_a, df_b)
        
        assert (1,) in result.key_only_in_a
        assert (4,) in result.key_only_in_b
        assert len(result.key_only_in_a) == 1
        assert len(result.key_only_in_b) == 1


class TestComplexScenarios:
    """Test complex, real-world scenarios"""
    
    def test_policy_coverage_comparison(self):
        """Test policy with multiple coverages"""
        df_a = pd.DataFrame({
            'Policy': ['P001', 'P001', 'P001', 'P002'],
            'Coverage': ['A', 'B', 'C', 'A'],
            'Premium': [100, 50, 25, 200],
            'Status': ['Active', 'Active', 'Active', 'Active']
        })
        df_b = pd.DataFrame({
            'Policy': ['P001', 'P001', 'P002', 'P002', 'P003'],
            'Coverage': ['A', 'C', 'A', 'B', 'A'],
            'Premium': [100, 25, 200, 75, 150],
            'Status': ['Active', 'Active', 'Active', 'Active', 'Active']
        })
        
        config = ComparisonConfig(key_columns=['Policy'])
        engine = ComparisonEngine(config)
        result = engine.compare(df_a, df_b)
        
        assert result.summary['keys_in_common'] == 2
        assert result.summary['keys_only_in_b'] == 1
        assert result.summary['added_row_count'] == 1  # P002 has extra coverage B
        assert result.summary['removed_row_count'] == 1  # P001 missing coverage B
    
    def test_multi_column_key_real_world(self):
        """Test real-world multi-column key scenario"""
        df_a = pd.DataFrame({
            'Account': ['ACC001', 'ACC001', 'ACC002'],
            'Date': ['2024-01-01', '2024-01-01', '2024-01-01'],
            'Amount': [1000.00, 500.00, 2000.00],
            'Description': ['Deposit', 'Fee', 'Transfer']
        })
        df_b = pd.DataFrame({
            'Account': ['ACC001', 'ACC001', 'ACC002'],
            'Date': ['2024-01-01', '2024-01-01', '2024-01-01'],
            'Amount': [1000.00, 500.00, 2100.00],  # Modified amount
            'Description': ['Deposit', 'Fee', 'Transfer']
        })
        
        config = ComparisonConfig(
            key_columns=['Account', 'Date'],
            alignment_method=AlignmentMethod.POSITION
        )
        engine = ComparisonEngine(config)
        result = engine.compare(df_a, df_b)
        
        assert result.summary['keys_in_common'] == 2
        assert result.summary['modified_count'] == 1  # One amount different
        assert result.summary['match_count'] == 2


if __name__ == '__main__':
    pytest.main([__file__, '-v'])
