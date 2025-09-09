#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Financial Statement Analysis

This script analyzes financial statements, calculates key metrics, and generates visualizations.

Author: Finance Team
Date: September 9, 2025
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
import os

class FinancialStatementAnalyzer:
    """Class for analyzing financial statements and generating insights."""
    
    def __init__(self, income_statement=None, balance_sheet=None, cash_flow=None):
        """
        Initialize with financial statements.
        
        Parameters:
        -----------
        income_statement : pandas.DataFrame, optional
            Income statement data with periods as columns and items as rows
        balance_sheet : pandas.DataFrame, optional
            Balance sheet data with periods as columns and items as rows
        cash_flow : pandas.DataFrame, optional
            Cash flow statement data with periods as columns and items as rows
        """
        self.income_statement = income_statement
        self.balance_sheet = balance_sheet
        self.cash_flow = cash_flow
        self.ratios = None
        
    def load_statements_from_excel(self, file_path):
        """
        Load financial statements from Excel file.
        
        Parameters:
        -----------
        file_path : str
            Path to Excel file containing financial statements
            
        Returns:
        --------
        bool
            True if loading was successful, False otherwise
        """
        try:
            # Read the Excel file
            self.income_statement = pd.read_excel(file_path, sheet_name='Income_Statement', index_col=0)
            self.balance_sheet = pd.read_excel(file_path, sheet_name='Balance_Sheet', index_col=0)
            self.cash_flow = pd.read_excel(file_path, sheet_name='Cash_Flow', index_col=0)
            
            print(f"Successfully loaded financial statements from {file_path}")
            return True
            
        except Exception as e:
            print(f"Error loading financial statements: {e}")
            return False
    
    def calculate_financial_ratios(self):
        """
        Calculate key financial ratios from the financial statements.
        
        Returns:
        --------
        pandas.DataFrame
            DataFrame containing calculated financial ratios
        """
        if self.income_statement is None or self.balance_sheet is None:
            print("Income statement and balance sheet are required to calculate ratios")
            return None
        
        # Create DataFrame to store ratios
        periods = self.income_statement.columns
        ratios = pd.DataFrame(index=['Current Ratio', 'Quick Ratio', 'Debt-to-Equity',
                                   'Gross Margin', 'Operating Margin', 'Net Margin',
                                   'Return on Assets', 'Return on Equity',
                                   'Inventory Turnover', 'Asset Turnover'],
                            columns=periods)
        
        try:
            # Liquidity Ratios
            # Current Ratio
            ratios.loc['Current Ratio'] = (
                self.balance_sheet.loc['Current Assets'] / 
                self.balance_sheet.loc['Current Liabilities']
            )
            
            # Quick Ratio
            ratios.loc['Quick Ratio'] = (
                (self.balance_sheet.loc['Current Assets'] - self.balance_sheet.loc['Inventory']) / 
                self.balance_sheet.loc['Current Liabilities']
            )
            
            # Leverage Ratios
            # Debt-to-Equity
            ratios.loc['Debt-to-Equity'] = (
                self.balance_sheet.loc['Total Liabilities'] / 
                self.balance_sheet.loc['Total Equity']
            )
            
            # Profitability Ratios
            # Gross Margin
            ratios.loc['Gross Margin'] = (
                self.income_statement.loc['Gross Profit'] / 
                self.income_statement.loc['Revenue']
            )
            
            # Operating Margin
            ratios.loc['Operating Margin'] = (
                self.income_statement.loc['Operating Income'] / 
                self.income_statement.loc['Revenue']
            )
            
            # Net Margin
            ratios.loc['Net Margin'] = (
                self.income_statement.loc['Net Income'] / 
                self.income_statement.loc['Revenue']
            )
            
            # Return on Assets
            ratios.loc['Return on Assets'] = (
                self.income_statement.loc['Net Income'] / 
                self.balance_sheet.loc['Total Assets']
            )
            
            # Return on Equity
            ratios.loc['Return on Equity'] = (
                self.income_statement.loc['Net Income'] / 
                self.balance_sheet.loc['Total Equity']
            )
            
            # Efficiency Ratios
            # Inventory Turnover
            ratios.loc['Inventory Turnover'] = (
                self.income_statement.loc['Cost of Goods Sold'] / 
                self.balance_sheet.loc['Inventory']
            )
            
            # Asset Turnover
            ratios.loc['Asset Turnover'] = (
                self.income_statement.loc['Revenue'] / 
                self.balance_sheet.loc['Total Assets']
            )
            
            self.ratios = ratios
            return ratios
            
        except Exception as e:
            print(f"Error calculating ratios: {e}")
            return None
    
    def plot_ratio_trends(self, ratio_category, figsize=(12, 6), save_path=None):
        """
        Plot trends for a category of financial ratios.
        
        Parameters:
        -----------
        ratio_category : str
            Category of ratios to plot ('liquidity', 'profitability', 'leverage', 'efficiency')
        figsize : tuple, optional
            Figure size (width, height) in inches
        save_path : str, optional
            Path to save the figure
            
        Returns:
        --------
        matplotlib.figure.Figure
            The created figure
        """
        if self.ratios is None:
            print("No ratios available. Call calculate_financial_ratios() first.")
            return None
        
        # Define ratio categories
        ratio_categories = {
            'liquidity': ['Current Ratio', 'Quick Ratio'],
            'profitability': ['Gross Margin', 'Operating Margin', 'Net Margin', 
                            'Return on Assets', 'Return on Equity'],
            'leverage': ['Debt-to-Equity'],
            'efficiency': ['Inventory Turnover', 'Asset Turnover']
        }
        
        if ratio_category.lower() not in ratio_categories:
            print(f"Invalid ratio category. Choose from: {list(ratio_categories.keys())}")
            return None
        
        # Get the ratios to plot
        ratios_to_plot = ratio_categories[ratio_category.lower()]
        
        # Create figure
        fig, ax = plt.subplots(figsize=figsize)
        
        # Plot each ratio
        for ratio in ratios_to_plot:
            if ratio in self.ratios.index:
                ax.plot(self.ratios.columns, self.ratios.loc[ratio], marker='o', label=ratio)
        
        # Set title and labels
        ax.set_title(f'{ratio_category.capitalize()} Ratios Trend', fontsize=14)
        ax.set_xlabel('Period', fontsize=12)
        ax.set_ylabel('Ratio Value', fontsize=12)
        
        # Add grid
        ax.grid(True, linestyle='--', alpha=0.7)
        
        # Add legend
        ax.legend()
        
        # Rotate x-axis labels for better readability
        plt.xticks(rotation=45)
        
        # Adjust layout
        plt.tight_layout()
        
        # Save figure if path is provided
        if save_path:
            plt.savefig(save_path, dpi=300, bbox_inches='tight')
        
        return fig
    
    def generate_summary_report(self, output_path=None):
        """
        Generate a summary report of financial statement analysis.
        
        Parameters:
        -----------
        output_path : str, optional
            Path to save the report
            
        Returns:
        --------
        str
            Summary report as string
        """
        if self.ratios is None:
            self.calculate_financial_ratios()
            
        if self.ratios is None:
            return "Unable to generate report without financial ratios."
        
        # Get the most recent period
        latest_period = self.ratios.columns[-1]
        previous_period = self.ratios.columns[-2] if len(self.ratios.columns) > 1 else None
        
        # Create summary report
        report = []
        report.append("# Financial Statement Analysis Summary Report")
        report.append(f"## Period: {latest_period}")
        report.append(f"Generated on: {datetime.now().strftime('%Y-%m-%d')}\n")
        
        report.append("## Key Financial Metrics\n")
        
        # Revenue and Profit
        report.append("### Revenue and Profit")
        revenue = self.income_statement.loc['Revenue', latest_period]
        net_income = self.income_statement.loc['Net Income', latest_period]
        report.append(f"- Revenue: ${revenue:,.2f}")
        report.append(f"- Net Income: ${net_income:,.2f}")
        report.append(f"- Net Margin: {self.ratios.loc['Net Margin', latest_period]:.2%}\n")
        
        # Liquidity
        report.append("### Liquidity")
        report.append(f"- Current Ratio: {self.ratios.loc['Current Ratio', latest_period]:.2f}")
        report.append(f"- Quick Ratio: {self.ratios.loc['Quick Ratio', latest_period]:.2f}\n")
        
        # Profitability
        report.append("### Profitability")
        report.append(f"- Return on Assets (ROA): {self.ratios.loc['Return on Assets', latest_period]:.2%}")
        report.append(f"- Return on Equity (ROE): {self.ratios.loc['Return on Equity', latest_period]:.2%}\n")
        
        # Period-over-period comparison if available
        if previous_period:
            report.append("## Period-over-Period Changes\n")
            
            # Revenue growth
            revenue_growth = (
                (self.income_statement.loc['Revenue', latest_period] / 
                 self.income_statement.loc['Revenue', previous_period]) - 1
            )
            report.append(f"- Revenue Growth: {revenue_growth:.2%}")
            
            # Net income growth
            net_income_growth = (
                (self.income_statement.loc['Net Income', latest_period] / 
                 self.income_statement.loc['Net Income', previous_period]) - 1
            )
            report.append(f"- Net Income Growth: {net_income_growth:.2%}")
            
            # Margin change
            margin_change = (
                self.ratios.loc['Net Margin', latest_period] - 
                self.ratios.loc['Net Margin', previous_period]
            )
            report.append(f"- Net Margin Change: {margin_change:.2%} points\n")
        
        # Insights and recommendations (placeholder)
        report.append("## Insights and Recommendations\n")
        report.append("1. Placeholder for automated insights based on financial analysis")
        report.append("2. Placeholder for recommendations based on trend analysis")
        report.append("3. Placeholder for risk assessment\n")
        
        # Convert report list to string
        report_text = "\n".join(report)
        
        # Save report if path is provided
        if output_path:
            with open(output_path, 'w') as f:
                f.write(report_text)
            print(f"Report saved to {output_path}")
        
        return report_text


def main():
    """Example usage of FinancialStatementAnalyzer."""
    print("Financial Statement Analyzer")
    print("---------------------------")
    
    # Create example data (in a real scenario, this would be loaded from files)
    example_data = {
        'income_statement': pd.DataFrame({
            '2023-Q1': [1000000, 600000, 400000, 250000, 150000],
            '2023-Q2': [1050000, 620000, 430000, 260000, 170000],
            '2023-Q3': [1100000, 640000, 460000, 270000, 190000],
            '2023-Q4': [1200000, 680000, 520000, 290000, 230000],
            '2024-Q1': [1250000, 700000, 550000, 300000, 250000],
        }, index=['Revenue', 'Cost of Goods Sold', 'Gross Profit', 'Operating Expenses', 'Net Income']),
        
        'balance_sheet': pd.DataFrame({
            '2023-Q1': [500000, 100000, 600000, 1500000, 2100000, 400000, 800000, 1200000, 900000],
            '2023-Q2': [520000, 110000, 630000, 1520000, 2150000, 410000, 820000, 1230000, 920000],
            '2023-Q3': [550000, 120000, 670000, 1550000, 2220000, 430000, 840000, 1270000, 950000],
            '2023-Q4': [600000, 130000, 730000, 1600000, 2330000, 450000, 880000, 1330000, 1000000],
            '2024-Q1': [650000, 140000, 790000, 1650000, 2440000, 470000, 920000, 1390000, 1050000],
        }, index=['Current Assets', 'Inventory', 'Total Current Assets', 'Non-Current Assets', 
                 'Total Assets', 'Current Liabilities', 'Long-term Liabilities', 'Total Liabilities', 'Total Equity'])
    }
    
    # Create analyzer with example data
    analyzer = FinancialStatementAnalyzer(
        income_statement=example_data['income_statement'],
        balance_sheet=example_data['balance_sheet']
    )
    
    # Calculate ratios
    ratios = analyzer.calculate_financial_ratios()
    
    print("\nThis script demonstrates financial statement analysis:")
    print("1. Loading financial statements from Excel files")
    print("2. Calculating key financial ratios")
    print("3. Plotting ratio trends over time")
    print("4. Generating a summary report with insights")
    
    print("\nTo use this script with your own data, provide file paths to your financial statements.")


if __name__ == "__main__":
    main()
