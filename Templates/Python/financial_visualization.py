#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Financial Visualization Template

This script provides a template for creating financial visualizations using matplotlib and seaborn.

Author: Finance Team
Date: September 9, 2025
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
import os

# Set the style for visualizations
plt.style.use('ggplot')
sns.set(style="whitegrid")


def create_time_series_plot(df, date_column, value_columns, title=None, figsize=(12, 6)):
    """
    Create a time series plot for financial data.
    
    Parameters:
    -----------
    df : pandas.DataFrame
        DataFrame containing financial data
    date_column : str
        Name of the column containing date values
    value_columns : list
        List of column names to plot
    title : str, optional
        Title for the plot
    figsize : tuple, optional
        Figure size (width, height) in inches
        
    Returns:
    --------
    matplotlib.figure.Figure
        The created figure
    """
    # Create figure and axis
    fig, ax = plt.subplots(figsize=figsize)
    
    # Plot each value column
    for col in value_columns:
        ax.plot(df[date_column], df[col], marker='o', linestyle='-', label=col)
    
    # Set title and labels
    if title:
        ax.set_title(title, fontsize=14)
    else:
        ax.set_title('Financial Time Series', fontsize=14)
    
    ax.set_xlabel('Date', fontsize=12)
    ax.set_ylabel('Value', fontsize=12)
    
    # Rotate x-axis labels for better readability
    plt.xticks(rotation=45)
    
    # Add legend
    ax.legend()
    
    # Adjust layout
    plt.tight_layout()
    
    return fig


def create_comparison_bar_chart(df, category_column, value_columns, title=None, figsize=(12, 6)):
    """
    Create a bar chart comparing multiple values across categories.
    
    Parameters:
    -----------
    df : pandas.DataFrame
        DataFrame containing financial data
    category_column : str
        Name of the column containing category values
    value_columns : list
        List of column names to compare
    title : str, optional
        Title for the plot
    figsize : tuple, optional
        Figure size (width, height) in inches
        
    Returns:
    --------
    matplotlib.figure.Figure
        The created figure
    """
    # Create figure and axis
    fig, ax = plt.subplots(figsize=figsize)
    
    # Set width of bars
    bar_width = 0.8 / len(value_columns)
    
    # Set positions of bars on X axis
    categories = df[category_column].unique()
    x = np.arange(len(categories))
    
    # Plot bars for each value column
    for i, col in enumerate(value_columns):
        positions = x + i * bar_width - (len(value_columns) - 1) * bar_width / 2
        values = [df[df[category_column] == cat][col].values[0] for cat in categories]
        ax.bar(positions, values, width=bar_width, label=col)
    
    # Set title and labels
    if title:
        ax.set_title(title, fontsize=14)
    else:
        ax.set_title('Financial Comparison', fontsize=14)
    
    ax.set_xlabel(category_column, fontsize=12)
    ax.set_ylabel('Value', fontsize=12)
    
    # Set x-tick labels
    ax.set_xticks(x)
    ax.set_xticklabels(categories, rotation=45, ha='right')
    
    # Add legend
    ax.legend()
    
    # Adjust layout
    plt.tight_layout()
    
    return fig


def create_financial_dashboard(df, date_column, metrics, figsize=(15, 10)):
    """
    Create a financial dashboard with multiple visualizations.
    
    Parameters:
    -----------
    df : pandas.DataFrame
        DataFrame containing financial data
    date_column : str
        Name of the column containing date values
    metrics : list
        List of metrics to include in the dashboard
    figsize : tuple, optional
        Figure size (width, height) in inches
        
    Returns:
    --------
    matplotlib.figure.Figure
        The created figure
    """
    # Create figure and axes
    fig, axes = plt.subplots(len(metrics), 1, figsize=figsize, sharex=True)
    
    # Create a plot for each metric
    for i, metric in enumerate(metrics):
        axes[i].plot(df[date_column], df[metric], marker='o', linestyle='-')
        axes[i].set_title(f'{metric} Over Time', fontsize=12)
        axes[i].set_ylabel(metric, fontsize=10)
        axes[i].grid(True, linestyle='--', alpha=0.7)
    
    # Set common x label
    axes[-1].set_xlabel('Date', fontsize=12)
    
    # Rotate x-axis labels for better readability
    plt.xticks(rotation=45)
    
    # Add overall title
    plt.suptitle('Financial Performance Dashboard', fontsize=16)
    
    # Adjust layout
    plt.tight_layout()
    plt.subplots_adjust(top=0.9)
    
    return fig


def main():
    """Main function to demonstrate template usage."""
    # Example usage
    print("Financial Visualization Template")
    print("-------------------------------")
    
    # Example data (to be replaced with actual data)
    dates = pd.date_range(start='2025-01-01', periods=12, freq='M')
    data = {
        'Date': dates,
        'Revenue': [100, 120, 140, 135, 150, 160, 170, 165, 180, 195, 200, 210],
        'Expenses': [80, 85, 90, 88, 95, 100, 105, 103, 110, 115, 120, 125],
        'Profit': [20, 35, 50, 47, 55, 60, 65, 62, 70, 80, 80, 85]
    }
    df = pd.DataFrame(data)
    
    print("\nThis template provides functions for:")
    print("1. Creating time series plots")
    print("2. Creating comparison bar charts")
    print("3. Building financial dashboards")
    print("\nReplace the example data with your actual financial data to use this template.")


if __name__ == "__main__":
    main()
