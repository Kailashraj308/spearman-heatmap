import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import numpy as np
from scipy.stats import spearmanr
import argparse
import os
import sys

def load_excel_data(file_path: str, sheet_name: str) -> pd.DataFrame:
    """
    Loads data from an Excel file and drops columns with missing values.
    """
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        return df.dropna(axis=1, how='any')  # Clean columns with NaN
    except Exception as e:
        print(f"Error loading Excel file: {e}")
        sys.exit(1)

def compute_spearman_corr(df: pd.DataFrame):
    """
    Computes Spearman correlation matrix and p-value matrix.
    """
    corr_matrix, pval_matrix = spearmanr(df)
    corr_df = pd.DataFrame(corr_matrix, columns=df.columns, index=df.columns)
    pval_df = pd.DataFrame(pval_matrix, columns=df.columns, index=df.columns)
    return corr_df, pval_df

def annotate_significance(corr_df: pd.DataFrame, pval_df: pd.DataFrame) -> pd.DataFrame:
    """
    Annotates correlation matrix with correlation values and significant p-values.
    """
    annot = corr_df.round(2).astype(str)
    for i in corr_df.index:
        for j in corr_df.columns:
            if i != j and pval_df.loc[i, j] < 0.05:
                annot.loc[i, j] += f"\n(p={pval_df.loc[i, j]:.2g})"
    return annot

def plot_heatmap(corr_df, annot_df, title, output_file):
    """
    Plots and saves a heatmap with correlation and p-value annotations.
    """
    plt.figure(figsize=(12, 10))
    sns.heatmap(corr_df, annot=annot_df, fmt='', cmap='coolwarm', center=0,
                linewidths=0.5, square=True)
    plt.title(title)
    plt.tight_layout()
    plt.savefig(output_file)
    print(f"Heatmap saved to: {os.path.abspath(output_file)}")
    plt.show()

def main():
    parser = argparse.ArgumentParser(description="Generate a Spearman correlation heatmap from an Excel sheet.")
    parser.add_argument("file", help="Path to the Excel file.")
    parser.add_argument("sheet", help="Sheet name in the Excel file.")
    parser.add_argument("--title", default="Spearman Correlation Matrix", help="Title of the heatmap.")
    parser.add_argument("--output", default="spearman_correlation_heatmap.png", help="Output image file name.")
    args = parser.parse_args()

    df = load_excel_data(args.file, args.sheet)
    corr_df, pval_df = compute_spearman_corr(df)
    annot_df = annotate_significance(corr_df, pval_df)
    plot_heatmap(corr_df, annot_df, args.title, args.output)

if __name__ == "__main__":
    main()
