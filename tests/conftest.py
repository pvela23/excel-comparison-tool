"""
Conftest for pytest configuration and shared fixtures
"""

import pytest
import pandas as pd
import tempfile
from pathlib import Path


@pytest.fixture
def sample_dataframe_a():
    """Create a sample DataFrame A"""
    return pd.DataFrame({
        'ID': [1, 2, 3, 4, 5],
        'Name': ['Alice', 'Bob', 'Charlie', 'David', 'Eve'],
        'Department': ['Sales', 'IT', 'HR', 'Finance', 'Sales'],
        'Salary': [50000, 60000, 55000, 65000, 51000],
        'Status': ['Active', 'Active', 'Inactive', 'Active', 'Active']
    })


@pytest.fixture
def sample_dataframe_b():
    """Create a sample DataFrame B with some differences"""
    return pd.DataFrame({
        'ID': [1, 2, 3, 4, 6],
        'Name': ['Alice', 'Bobby', 'Charlie', 'David', 'Frank'],
        'Department': ['Sales', 'IT', 'HR', 'Finance', 'IT'],
        'Salary': [52000, 60000, 55000, 67000, 62000],  # Modified some values
        'Status': ['Active', 'Active', 'Inactive', 'Active', 'Active']
    })


@pytest.fixture
def sample_policy_dataframe_a():
    """Create a sample policy DataFrame with multiple rows per key"""
    return pd.DataFrame({
        'PolicyNumber': ['POL001', 'POL001', 'POL001', 'POL002', 'POL002'],
        'CoverageType': ['Liability', 'Property', 'Medical', 'Liability', 'Property'],
        'Premium': [500, 1000, 200, 600, 800],
        'Limit': [100000, 500000, 10000, 100000, 300000]
    })


@pytest.fixture
def sample_policy_dataframe_b():
    """Create a sample policy DataFrame with slight differences"""
    return pd.DataFrame({
        'PolicyNumber': ['POL001', 'POL001', 'POL001', 'POL002', 'POL002', 'POL003'],
        'CoverageType': ['Liability', 'Property', 'Medical', 'Liability', 'Property', 'Liability'],
        'Premium': [525, 1000, 200, 600, 800, 450],  # Modified first value, added last
        'Limit': [100000, 500000, 10000, 100000, 300000, 75000]
    })


@pytest.fixture
def temp_excel_file():
    """Create a temporary Excel file"""
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
        temp_path = f.name
    
    yield temp_path
    
    # Cleanup
    if Path(temp_path).exists():
        Path(temp_path).unlink()


@pytest.fixture
def temp_report_file():
    """Create a temporary report Excel file"""
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
        temp_path = f.name
    
    yield temp_path
    
    # Cleanup
    if Path(temp_path).exists():
        Path(temp_path).unlink()


@pytest.fixture(autouse=True)
def reset_random_state():
    """Reset random state before each test"""
    import random
    import numpy as np
    random.seed(42)
    np.random.seed(42)
