"""
Test utilities and helper functions
"""

import pandas as pd
import numpy as np
from typing import Dict, Any
from src.core import RowStatus


class TestDataGenerator:
    """Helper class for generating test data"""
    
    @staticmethod
    def create_simple_dataframe(n_rows: int = 10) -> pd.DataFrame:
        """
        Create a simple test DataFrame
        
        Args:
            n_rows: Number of rows to generate
            
        Returns:
            DataFrame with ID, Name, and Value columns
        """
        return pd.DataFrame({
            'ID': range(1, n_rows + 1),
            'Name': [f'Person_{i}' for i in range(1, n_rows + 1)],
            'Value': np.random.randint(100, 1000, n_rows)
        })
    
    @staticmethod
    def create_dataframe_with_nan(n_rows: int = 5) -> pd.DataFrame:
        """
        Create a DataFrame with some NaN values
        
        Args:
            n_rows: Number of rows to generate
            
        Returns:
            DataFrame with some NaN values
        """
        df = TestDataGenerator.create_simple_dataframe(n_rows)
        # Add some NaN values
        df.loc[1, 'Value'] = np.nan
        df.loc[3, 'Name'] = np.nan
        return df
    
    @staticmethod
    def create_policy_dataframe(n_policies: int = 5, coverages_per_policy: int = 2) -> pd.DataFrame:
        """
        Create a DataFrame with policy and coverage data
        
        Args:
            n_policies: Number of policies to create
            coverages_per_policy: Average coverages per policy
            
        Returns:
            DataFrame with policy structure
        """
        rows = []
        for p in range(1, n_policies + 1):
            for c in range(1, coverages_per_policy + 1):
                rows.append({
                    'PolicyNumber': f'P{p:03d}',
                    'CoverageType': f'Type_{c}',
                    'Premium': p * 100 + c * 10,
                    'Limit': (p * 100000) + (c * 10000)
                })
        return pd.DataFrame(rows)
    
    @staticmethod
    def create_transaction_dataframe(n_transactions: int = 10) -> pd.DataFrame:
        """
        Create a DataFrame with transaction data
        
        Args:
            n_transactions: Number of transactions to create
            
        Returns:
            DataFrame with transaction structure
        """
        return pd.DataFrame({
            'TransactionID': range(1, n_transactions + 1),
            'Account': [f'ACC{(i % 3) + 1:03d}' for i in range(n_transactions)],
            'Date': pd.date_range('2024-01-01', periods=n_transactions),
            'Amount': np.random.uniform(10, 10000, n_transactions),
            'Type': np.random.choice(['Deposit', 'Withdrawal', 'Fee'], n_transactions)
        })


class TestResultValidator:
    """Helper class for validating test results"""
    
    @staticmethod
    def validate_summary_structure(summary: Dict[str, Any]) -> bool:
        """
        Validate that summary has all required keys
        
        Args:
            summary: Summary dictionary to validate
            
        Returns:
            True if valid, False otherwise
        """
        required_keys = [
            'total_unique_keys_a',
            'total_unique_keys_b',
            'keys_in_common',
            'keys_only_in_a',
            'keys_only_in_b',
            'total_rows_compared',
            'match_count',
            'modified_count',
            'added_row_count',
            'removed_row_count',
            'new_key_count',
            'removed_key_count',
        ]
        return all(key in summary for key in required_keys)
    
    @staticmethod
    def validate_row_status_values(aligned_data: pd.DataFrame) -> bool:
        """
        Validate that all status values in aligned data are valid
        
        Args:
            aligned_data: DataFrame with status column
            
        Returns:
            True if all statuses are valid, False otherwise
        """
        if 'status' not in aligned_data.columns:
            return False
        
        valid_statuses = [status.value for status in RowStatus]
        return all(status in valid_statuses for status in aligned_data['status'])
    
    @staticmethod
    def count_rows_by_status(aligned_data: pd.DataFrame) -> Dict[str, int]:
        """
        Count rows by their status
        
        Args:
            aligned_data: DataFrame with status column
            
        Returns:
            Dictionary with status counts
        """
        if 'status' not in aligned_data.columns:
            return {}
        
        return aligned_data['status'].value_counts().to_dict()


class TestFileGenerator:
    """Helper class for generating test files"""
    
    @staticmethod
    def save_test_excel(df: pd.DataFrame, file_path: str, sheet_name: str = 'Sheet1'):
        """
        Save a DataFrame as an Excel file for testing
        
        Args:
            df: DataFrame to save
            file_path: Path where to save the file
            sheet_name: Name of the Excel sheet
        """
        df.to_excel(file_path, sheet_name=sheet_name, index=False)
    
    @staticmethod
    def create_test_excel_pair(df_a: pd.DataFrame, df_b: pd.DataFrame, 
                              file_a_path: str, file_b_path: str):
        """
        Create a pair of Excel files for testing
        
        Args:
            df_a: DataFrame for file A
            df_b: DataFrame for file B
            file_a_path: Path for file A
            file_b_path: Path for file B
        """
        TestFileGenerator.save_test_excel(df_a, file_a_path, 'Data')
        TestFileGenerator.save_test_excel(df_b, file_b_path, 'Data')


if __name__ == '__main__':
    # Example usage
    generator = TestDataGenerator()
    df = generator.create_simple_dataframe(5)
    print(df)
