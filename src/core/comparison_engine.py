# Comparison engine

"""
Core Comparison Engine for Excel Comparison Tool
Handles key-based comparison with multi-row support
"""

import pandas as pd # type: ignore
import numpy as np # type: ignore
from typing import List, Dict, Tuple, Optional, Any
from dataclasses import dataclass, field
from enum import Enum


class RowStatus(Enum):
    """Status types for compared rows"""
    MATCH = "MATCH"
    MODIFIED = "MODIFIED"
    ADDED_ROW = "ADDED_ROW"
    REMOVED_ROW = "REMOVED_ROW"
    NEW_KEY = "NEW_KEY"
    REMOVED_KEY = "REMOVED_KEY"
    HEURISTIC_MATCH = "HEURISTIC_MATCH"


class AlignmentMethod(Enum):
    """Methods for aligning rows within key groups"""
    POSITION = "position"
    SECONDARY_SORT = "secondary_sort"
    BEST_MATCH = "best_match"


@dataclass
class ComparisonConfig:
    """Configuration for comparison operation"""
    key_columns: List[str]
    alignment_method: AlignmentMethod = AlignmentMethod.POSITION
    secondary_sort_column: Optional[str] = None
    case_sensitive: bool = False
    trim_whitespace: bool = True
    compare_formulas: bool = False  # Post-MVP
   

@dataclass
class ComparisonResult:
    """Results of a comparison operation"""
    summary: Dict[str, Any] = field(default_factory=dict)
    aligned_data: pd.DataFrame = field(default_factory=pd.DataFrame)
    key_only_in_a: List[Tuple] = field(default_factory=list)
    key_only_in_b: List[Tuple] = field(default_factory=list)
    comparison_metadata: Dict[str, Any] = field(default_factory=dict)


class ComparisonEngine:
    """
    Core engine for comparing Excel files with multi-row key support
    """
   
    def __init__(self, config: ComparisonConfig):
        self.config = config
       
    def compare(self, df_a: pd.DataFrame, df_b: pd.DataFrame) -> ComparisonResult:
        """
        Main comparison method
       
        Args:
            df_a: DataFrame from File A
            df_b: DataFrame from File B
           
        Returns:
            ComparisonResult with aligned data and summary
        """
        # Validate inputs
        self._validate_dataframes(df_a, df_b)
       
        # Normalize data
        df_a = self._normalize_dataframe(df_a.copy())
        df_b = self._normalize_dataframe(df_b.copy())
       
        # Get unique keys from both files
        keys_a = self._get_unique_keys(df_a)
        keys_b = self._get_unique_keys(df_b)
       
        # Identify key differences
        keys_only_a = keys_a - keys_b
        keys_only_b = keys_b - keys_a
        keys_common = keys_a & keys_b
       
        # Process each key group
        aligned_rows = []
       
        # Process common keys (main comparison logic)
        for key in sorted(keys_common):
            key_results = self._compare_key_group(key, df_a, df_b)
            aligned_rows.extend(key_results)
       
        # Process keys only in A (removed keys)
        for key in sorted(keys_only_a):
            rows_a = self._get_rows_for_key(df_a, key)
            for _, row in rows_a.iterrows():
                aligned_rows.append(self._create_aligned_row(
                    key, row, None, RowStatus.REMOVED_KEY
                ))
       
        # Process keys only in B (new keys)
        for key in sorted(keys_only_b):
            rows_b = self._get_rows_for_key(df_b, key)
            for _, row in rows_b.iterrows():
                aligned_rows.append(self._create_aligned_row(
                    key, None, row, RowStatus.NEW_KEY
                ))
       
        # Create aligned DataFrame
        aligned_df = pd.DataFrame(aligned_rows)
       
        # Generate summary statistics
        summary = self._generate_summary(
            aligned_df, len(keys_a), len(keys_b), len(keys_common)
        )
       
        return ComparisonResult(
            summary=summary,
            aligned_data=aligned_df,
            key_only_in_a=sorted(list(keys_only_a)),
            key_only_in_b=sorted(list(keys_only_b)),
            comparison_metadata={
                'config': self.config,
                'total_keys_compared': len(keys_common),
                'total_rows_a': len(df_a),
                'total_rows_b': len(df_b)
            }
        )
   
    def _validate_dataframes(self, df_a: pd.DataFrame, df_b: pd.DataFrame):
        """Validate that key columns exist in both dataframes"""
        for key_col in self.config.key_columns:
            if key_col not in df_a.columns:
                raise KeyError(f"Key column '{key_col}' not found in File A")
            if key_col not in df_b.columns:
                raise KeyError(f"Key column '{key_col}' not found in File B")
   
    def _normalize_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        """Apply normalization rules to DataFrame"""
        if self.config.trim_whitespace:
            for col in df.select_dtypes(include=['object']).columns:
                df[col] = df[col].astype(str).str.strip()
       
        if not self.config.case_sensitive:
            for col in df.select_dtypes(include=['object']).columns:
                df[col] = df[col].astype(str).str.lower()
       
        return df
   
    def _get_unique_keys(self, df: pd.DataFrame) -> set:
        """Extract unique key tuples from DataFrame"""
        if len(self.config.key_columns) == 1:
            keys = df[self.config.key_columns[0]].unique()
            return set((k,) for k in keys)
        else:
            keys = df[self.config.key_columns].drop_duplicates()
            return set(tuple(row) for row in keys.values)
   
    def _get_rows_for_key(self, df: pd.DataFrame, key: Tuple) -> pd.DataFrame:
        """Get all rows matching a specific key"""
        mask = pd.Series([True] * len(df))
        for i, col in enumerate(self.config.key_columns):
            mask &= (df[col] == key[i])
        return df[mask].copy()
   
    def _compare_key_group(
        self,
        key: Tuple,
        df_a: pd.DataFrame,
        df_b: pd.DataFrame
    ) -> List[Dict]:
        """
        Compare all rows within a single key group
       
        This is the core logic that handles multiple rows per key
        """
        rows_a = self._get_rows_for_key(df_a, key)
        rows_b = self._get_rows_for_key(df_b, key)
       
        # Apply alignment method
        if self.config.alignment_method == AlignmentMethod.SECONDARY_SORT:
            rows_a = self._sort_by_secondary(rows_a)
            rows_b = self._sort_by_secondary(rows_b)
       
        # Convert to lists for easier indexing
        rows_a_list = [row for _, row in rows_a.iterrows()]
        rows_b_list = [row for _, row in rows_b.iterrows()]
       
        aligned_results = []
       
        # Position-based alignment
        max_rows = max(len(rows_a_list), len(rows_b_list))
       
        for i in range(max_rows):
            row_a = rows_a_list[i] if i < len(rows_a_list) else None
            row_b = rows_b_list[i] if i < len(rows_b_list) else None
           
            if row_a is not None and row_b is not None:
                # Both rows exist - compare them
                status = self._compare_rows(row_a, row_b)
                aligned_results.append(
                    self._create_aligned_row(key, row_a, row_b, status)
                )
            elif row_a is not None:
                # Row only in A - removed
                aligned_results.append(
                    self._create_aligned_row(key, row_a, None, RowStatus.REMOVED_ROW)
                )
            else:
                # Row only in B - added
                aligned_results.append(
                    self._create_aligned_row(key, None, row_b, RowStatus.ADDED_ROW)
                )
       
        return aligned_results
   
    def _sort_by_secondary(self, df: pd.DataFrame) -> pd.DataFrame:
        """Sort rows by secondary sort column"""
        if self.config.secondary_sort_column and self.config.secondary_sort_column in df.columns:
            return df.sort_values(by=self.config.secondary_sort_column)
        return df
   
    def _compare_rows(self, row_a: pd.Series, row_b: pd.Series) -> RowStatus:
        """
        Compare two rows cell-by-cell (excluding key columns)
        """
        # Get non-key columns
        non_key_cols = [col for col in row_a.index if col not in self.config.key_columns]
       
        for col in non_key_cols:
            if col not in row_b.index:
                continue
           
            val_a = row_a[col]
            val_b = row_b[col]
           
            # Handle NaN/None
            if pd.isna(val_a) and pd.isna(val_b):
                continue
            if pd.isna(val_a) or pd.isna(val_b):
                return RowStatus.MODIFIED
           
            # Compare values
            if val_a != val_b:
                return RowStatus.MODIFIED
       
        return RowStatus.MATCH
   
    def _create_aligned_row(
        self,
        key: Tuple,
        row_a: Optional[pd.Series],
        row_b: Optional[pd.Series],
        status: RowStatus
    ) -> Dict:
        """Create a single aligned row for output"""
        result = {}
       
        # Add key columns
        for i, col_name in enumerate(self.config.key_columns):
            result[f'key_{col_name}'] = key[i]
       
        # Add File A columns
        if row_a is not None:
            for col in row_a.index:
                if col not in self.config.key_columns:
                    result[f'A_{col}'] = row_a[col]
       
        # Add status
        result['status'] = status.value
       
        # Add File B columns
        if row_b is not None:
            for col in row_b.index:
                if col not in self.config.key_columns:
                    result[f'B_{col}'] = row_b[col]
       
        # Add changed cells info (for MODIFIED rows)
        if status == RowStatus.MODIFIED and row_a is not None and row_b is not None:
            changed_cells = []
            for col in row_a.index:
                if col not in self.config.key_columns and col in row_b.index:
                    if not self._values_equal(row_a[col], row_b[col]):
                        changed_cells.append(col)
            result['changed_cells'] = ', '.join(changed_cells)
       
        return result
   
    def _values_equal(self, val_a, val_b) -> bool:
        """Compare two values handling NaN"""
        if pd.isna(val_a) and pd.isna(val_b):
            return True
        if pd.isna(val_a) or pd.isna(val_b):
            return False
        return val_a == val_b
   
    def _generate_summary(
        self,
        aligned_df: pd.DataFrame,
        total_keys_a: int,
        total_keys_b: int,
        keys_common: int
    ) -> Dict[str, Any]:
        """Generate summary statistics"""
        status_counts = aligned_df['status'].value_counts().to_dict()
       
        return {
            'total_unique_keys_a': total_keys_a,
            'total_unique_keys_b': total_keys_b,
            'keys_in_common': keys_common,
            'keys_only_in_a': total_keys_a - keys_common,
            'keys_only_in_b': total_keys_b - keys_common,
            'total_rows_compared': len(aligned_df),
            'match_count': status_counts.get(RowStatus.MATCH.value, 0),
            'modified_count': status_counts.get(RowStatus.MODIFIED.value, 0),
            'added_row_count': status_counts.get(RowStatus.ADDED_ROW.value, 0),
            'removed_row_count': status_counts.get(RowStatus.REMOVED_ROW.value, 0),
            'new_key_count': status_counts.get(RowStatus.NEW_KEY.value, 0),
            'removed_key_count': status_counts.get(RowStatus.REMOVED_KEY.value, 0),
        }

