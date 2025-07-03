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
    Loads data from an Excel file, selects numeric columns, and drops columns with missing values.
    """
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        df = df.select_dtypes(include=[np.number])  # Keep only numeric columns
        df = df.dropna(axis=1, how='any')           # Drop columns with any NaN values
        if df.empty:
            print("No usable numeric data available after cleaning.")
            sys.exit(1)
        return df
    except Exception as e:
        print(f"Error loading Excel file: {e}")
        sys.exit(1)

def compute_spearman_corr(df: pd.DataFrame):
    """
    Computes Spearman correlation and p-value matrices pairwise.
    """
    cols = df.columns
    corr_df = pd.DataFrame(index=cols, columns=cols, dtype=float)
    pval_df = pd.DataFrame(index=cols, columns=cols, dtype=float)

    for col1 in cols:
        for col2 in cols:
            rho, pval = spearmanr(df[col1], df[col2])
            corr_df.loc[col1, col2] = rho
            pval_df.loc[col1, col2] = pval

    return corr_df, pval_df

def annotate_significance(corr_df: pd.DataFrame, pval_df: pd.DataFrame) -> pd.DataFrame:
    """
    Annotates the correlation matrix with correlation values and significant p-values.
    """
    annot = corr_df.round(2).astype(str)
    for i in corr_df.index:
        for j in corr_df.columns:
            if i != j and pval_df.loc[i, j] < 0.05:
                annot.loc[i, j] += f"\n(p={pval_df.loc[i, j]:.2g})"
    return annot

def plot_heatmap(corr_df, annot_df, title, output_file):
    """
    Plots and saves a heatmap with annotated correlation values.
    """
    plt.figure(figsize=(12, 10))
    sns.heatmap(corr_df, annot=annot_df, fmt='', cmap='coolwarm', center=0,
                linewidths=0.5, square=True, cbar_kws={'shrink': 0.8})
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
    parser.add_argument("--export_csv", action="store_true", help="Export correlation and p-value matrices as CSV.")
    args = parser.parse_args()

    df = load_excel_data(args.file, args.sheet)
    corr_df, pval_df = compute_spearman_corr(df)
    annot_df = annotate_significance(corr_df, pval_df)
    plot_heatmap(corr_df, annot_df, args.title, args.output)

    if args.export_csv:
        corr_df.to_csv("correlation_matrix.csv")
        pval_df.to_csv("pvalue_matrix.csv")
        print("CSV files exported: correlation_matrix.csv, pvalue_matrix.csv")

if __name__ == "__main__":
    main()
