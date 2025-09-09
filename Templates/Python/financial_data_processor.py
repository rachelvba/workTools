#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Financial Data Processing Template

This script provides a template for processing financial data using pandas.

Author: Finance Team
Date: September 9, 2025
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from datetime import datetime
import os

def load_financial_data(file_path, sheet_name=None):
    """
    Load financial data from an Excel or CSV file.
    
    Parameters:
    -----------
    file_path : str
        Path to the data file
    sheet_name : str, optional
        Name of the Excel sheet to load (None for CSV files)
        
    Returns:
    --------
    pandas.DataFrame
        Loaded financial data
    """
    try:
        # Check file extension
        _, ext = os.path.splitext(file_path)
        
        if ext.lower() in ['.xlsx', '.xls']:
            # Excel file
            df = pd.read_excel(file_path, sheet_name=sheet_name)
        elif ext.lower() == '.csv':
            # CSV file
            df = pd.read_csv(file_path)
        else:
            raise ValueError(f"Unsupported file format: {ext}")
        
        print(f"Successfully loaded data with {df.shape[0]} rows and {df.shape[1]} columns")
        return df
    
    except Exception as e:
        print(f"Error loading data: {e}")
        return None


def clean_financial_data(df):
    """
    Clean and prepare financial data for analysis.
    
    Parameters:
    -----------
    df : pandas.DataFrame
        DataFrame containing financial data
        
    Returns:
    --------
    pandas.DataFrame
        Cleaned DataFrame
    """
    # Create a copy to avoid modifying the original
    df_clean = df.copy()
    
    # Convert date columns
    date_cols = [col for col in df_clean.columns if 'date' in col.lower()]
    for col in date_cols:
        try:
            df_clean[col] = pd.to_datetime(df_clean[col])
        except:
            print(f"Could not convert column {col} to datetime")
    
    # Handle missing values
    numeric_cols = df_clean.select_dtypes(include=['number']).columns
    for col in numeric_cols:
        # Fill missing values with median
        median_val = df_clean[col].median()
        df_clean[col].fillna(median_val, inplace=True)
    
    # Drop rows with all missing values
    df_clean.dropna(how='all', inplace=True)
    
    # Remove duplicates
    df_clean.drop_duplicates(inplace=True)
    
    return df_clean


def calculate_financial_metrics(df):
    """
    Calculate common financial metrics and add them to the DataFrame.
    
    Parameters:
    -----------
    df : pandas.DataFrame
        DataFrame containing financial data
        
    Returns:
    --------
    pandas.DataFrame
        DataFrame with additional financial metrics
    """
    # Create a copy to avoid modifying the original
    df_metrics = df.copy()
    
    # Example: Calculate growth rates if 'Revenue' column exists
    if 'Revenue' in df_metrics.columns:
        df_metrics['Revenue_Growth'] = df_metrics['Revenue'].pct_change() * 100
    
    # Example: Calculate profit margin if 'Revenue' and 'Profit' columns exist
    if all(col in df_metrics.columns for col in ['Revenue', 'Profit']):
        df_metrics['Profit_Margin'] = (df_metrics['Profit'] / df_metrics['Revenue']) * 100
    
    # Example: Calculate moving averages
    if 'Revenue' in df_metrics.columns:
        df_metrics['Revenue_3MA'] = df_metrics['Revenue'].rolling(window=3).mean()
    
    return df_metrics


def main():
    """Main function to demonstrate template usage."""
    # Example usage
    print("Financial Data Processing Template")
    print("----------------------------------")
    
    # Define file path (to be replaced with actual file path)
    file_path = "path/to/your/financial_data.xlsx"
    
    # Example workflow
    print("\nThis template would typically:")
    print("1. Load financial data from Excel or CSV")
    print("2. Clean the data (handle missing values, date formatting)")
    print("3. Calculate financial metrics")
    print("4. Perform analysis or create visualizations")
    print("\nReplace 'file_path' with your actual data file to use this template.")


if __name__ == "__main__":
    main()
