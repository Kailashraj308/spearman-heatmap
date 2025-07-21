# -*- coding: utf-8 -*-
"""
Interactive Spearman Correlation Analysis and Heatmap Generation

This script interactively prompts the user for a file path, a sheet name,
and a title, then performs a Spearman rank-order correlation analysis,
generates a heatmap, and saves the plot to a file. (Reccomended to run in google colab)

Author: Kailash Rajpurohit
License: MIT
"""

import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import numpy as np
from scipy.stats import spearmanr
import os
import re

def load_and_clean_data(file_path, sheet_name):
    """
    Loads data from a specific sheet of an Excel file and cleans it.

    Args:
        file_path (str): The path to the input Excel file.
        sheet_name (str): The name of the sheet to load.

    Returns:
        pd.DataFrame: A cleaned pandas DataFrame with no missing values,
                      or None if the file/sheet cannot be loaded.
    """
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        print(f"Successfully loaded sheet: '{sheet_name}'")
        
        # Drop rows with any missing values to ensure calculations are consistent.
        df_clean = df.dropna(axis=0, how='any')
        
        print(f"Data cleaned. Original rows: {len(df)}, Cleaned rows: {len(df_clean)}")
        return df_clean

    except FileNotFoundError:
        print(f"Error: The file was not found at '{file_path}'")
        return None
    except Exception as e:
        print(f"An error occurred while loading the data: {e}")
        return None

def calculate_spearman_correlation(df):
    """
    Calculates the Spearman correlation matrix and p-values for a DataFrame.

    Args:
        df (pd.DataFrame): The input DataFrame (must contain only numeric data).

    Returns:
        tuple: A tuple containing the correlation DataFrame and the p-value DataFrame.
    """
    if df.shape[1] < 2:
        print("Error: Not enough numeric columns to perform correlation.")
        return None, None
        
    corr_matrix, pval_matrix = spearmanr(df)
    corr_df = pd.DataFrame(corr_matrix, columns=df.columns, index=df.columns)
    pval_df = pd.DataFrame(pval_matrix, columns=df.columns, index=df.columns)
    
    return corr_df, pval_df

def generate_heatmap(corr_df, pval_df, title, output_path):
    """
    Generates and saves a heatmap from correlation and p-value data.

    Args:
        corr_df (pd.DataFrame): DataFrame of correlation coefficients.
        pval_df (pd.DataFrame): DataFrame of p-values.
        title (str): The title for the heatmap plot.
        output_path (str): The path to save the generated heatmap image.
    """
    # Create annotations: format correlation coefficient and add p-value if significant
    annot = corr_df.round(2).astype(str)
    for i in corr_df.index:
        for j in corr_df.columns:
            if i != j and pval_df.loc[i, j] < 0.05:
                annot.loc[i, j] += f"\n(p={pval_df.loc[i, j]:.2g})"

    # Plotting
    plt.figure(figsize=(14, 12))
    sns.heatmap(corr_df, annot=annot, fmt='', cmap='coolwarm', center=0, linewidths=0.5, square=True)
    plt.title(title, fontsize=16)
    plt.tight_layout(pad=3.0)
    
    try:
        plt.savefig(output_path, dpi=300)
        print(f"Heatmap successfully saved to: '{output_path}'")
    except Exception as e:
        print(f"Error saving heatmap: {e}")
        
    plt.close()

def main():
    """
    Main function to interactively get user input and run the analysis.
    """
    print("--- Interactive Spearman Correlation Analysis ---")
    
    # 1. Get user input
    file_path = input("Enter the full path to your Excel file: ").strip()
    sheet_name = input("Enter the exact name of the sheet to analyze: ").strip()
    heatmap_title = input("Enter the desired title for the heatmap: ").strip()
    
    # --- Workflow ---
    # 2. Load and clean data
    df_clean = load_and_clean_data(file_path, sheet_name)
    
    if df_clean is None or df_clean.empty:
        print("Exiting: Data could not be loaded or is empty after cleaning.")
        return

    # 3. Calculate correlation
    corr_df, pval_df = calculate_spearman_correlation(df_clean)
    
    if corr_df is None:
        print("Exiting: Correlation could not be calculated.")
        return

    # 4. Create output directory if it doesn't exist
    output_dir = "output"
    os.makedirs(output_dir, exist_ok=True)
    
    # 5. Generate a safe filename from the user's title
    # Remove special characters and replace spaces with underscores
    safe_filename = re.sub(r'[^a-zA-Z0-9_\- ]', '', heatmap_title).replace(' ', '_')
    output_filename = f"{safe_filename}.png"
    output_path = os.path.join(output_dir, output_filename)
    
    # 6. Generate and save the heatmap
    generate_heatmap(corr_df, pval_df, heatmap_title, output_path)

if __name__ == "__main__":
    main()

